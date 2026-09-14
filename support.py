import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Alignment

# 다중 시트 엑셀 다운로드를 위한 변환 함수 (줄바꿈 서식 및 열 너비 조정)
def to_excel_multi_sheet(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            
            worksheet = writer.sheets[sheet_name]
            
            # 모든 셀에 텍스트 줄바꿈 및 상단 정렬 적용
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                    
            # 열 너비 자동 조절 (최대 80 제한)
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            if len(line) > max_length:
                                max_length = len(line)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 80)
                worksheet.column_dimensions[column].width = adjusted_width

    return output.getvalue()

# 일시 양식을 '0000년 00월 00일 00:00' 포맷으로 표준화
def standardize_date(date_raw):
    date_str = re.sub(r'\s+', ' ', str(date_raw)).strip()
    nums = re.findall(r'\d+', date_str)
    
    if len(nums) >= 4:
        year = nums[0]
        month = nums[1].zfill(2)
        day = nums[2].zfill(2)
        hour = int(nums[3])
        minute = int(nums[4]) if len(nums) >= 5 else 0
        
        if '오후' in date_str and hour < 12:
            hour += 12
        elif '오전' in date_str and hour == 12:
            hour = 0
            
        return f"{year}년 {month}월 {day}일 {hour:02d}:{minute:02d}"
    return date_str

# 상단 헤더 영역에서 학교명, 학교급, 담당 장학사를 추출하는 헬퍼 함수
def extract_header_info(df, filename=""):
    raw_school_name = "학교명 미상"
    supervisor_name = "장학사 미상"
    school_level = ""
    level_sort_val = 3

    # 상단 12개 행 탐색
    for idx in range(min(12, len(df))):
        val = str(df.iloc[idx, 0]).strip()
        if not val or val.startswith("【") or val.startswith("※") or val.startswith("작성") or val.startswith("("):
            continue
        if val == "순":
            break
            
        # 담당 장학사 추출
        if "장학사" in val:
            sup_match = re.search(r'(?:학교담당장학사|담당장학사|장학사)\s*[:：]?\s*([가-힣]{2,4})', val)
            if sup_match and supervisor_name == "장학사 미상":
                supervisor_name = sup_match.group(1).strip()
        
        # 학교명 추출 (안내 문구 및 직책 행 제외)
        if ("학교" in val or "중" in val or "고" in val) and "장학사" not in val:
            if raw_school_name == "학교명 미상":
                raw_school_name = val

    # 파일명 기반 보정
    if raw_school_name == "학교명 미상" and filename:
        parts = filename.replace(".xlsx", "").replace(".csv", "").split("_")
        if len(parts) > 1:
            raw_school_name = parts[1]

    # 학교급 판별 및 정렬 가중치 부여
    if "중학교" in raw_school_name or filename.startswith("중_"):
        school_level = "중"
        level_sort_val = 1
    elif "고등학교" in raw_school_name or filename.startswith("고_"):
        school_level = "고"
        level_sort_val = 2

    return raw_school_name, school_level, level_sort_val, supervisor_name

# 페이지 설정
st.set_page_config(page_title="지원장학 요청서 자동 분석기", layout="wide")
st.title("📊 지원장학 요청서 자동 분석 및 유목화 웹앱")
st.markdown("지원장학 요청서를 업로드하면 지정된 셀에서 데이터를 추출하고, 조건에 맞게 정렬하여 엑셀 파일로 추출합니다.")

uploaded_files = st.file_uploader("장학 요청서 파일(Excel 또는 CSV)을 업로드하세요.", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("분석 시작") and uploaded_files:
    schedule_list = []
    issue_list = []
    request_list = []
    categorized_list = []
    
    # 세부 키워드 보강 (급식, 기자재, 진학, 출결, 심리 등)
    categories = {
        "시설 및 환경 개선": ["방송", "공사", "누수", "노후", "수리", "교체", "안전", "공간", "장비", "기자재", "급식", "냉난방", "환경"],
        "예산 및 행정 지원": ["예산", "지원금", "품의", "결제", "계약", "행정", "인력", "채용", "강사", "회계"],
        "교육과정 및 학사 운영": ["교과", "학점제", "평가", "성적", "교육과정", "자유학기", "수업", "디지털", "코딩", "디벗", "진학", "학사", "체험학습"],
        "생활지도 및 학생 지원": ["폭력", "학폭", "상담", "정서", "위기", "징계", "출결", "다문화", "심리", "생명존중", "자해", "자살"],
        "기타(미분류)": []
    }

    dept_mapping = {
        "시설 및 환경 개선": "교육재정상담과(또는 교육시설과)",
        "예산 및 행정 지원": "행정지원국(행정지원과)",
        "교육과정 및 학사 운영": "교육지원국(중등교육과)",
        "생활지도 및 학생 지원": "학생학부모지원센터(또는 학교통합지원센터)",
        "기타(미분류)": "관련 부서 확인 필요"
    }

    st.success(f"총 {len(uploaded_files)}개의 파일을 분석합니다...")
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, header=None, dtype=str).fillna("")
            else:
                df = pd.read_excel(file, header=None, dtype=str).fillna("")

            # 1. 학교 기본정보 추출
            raw_school_name, school_level, level_sort_val, supervisor_name = extract_header_info(df, file.name)

            # 2. 표 헤더('구분', '내용') 열 위치 동적 감지
            header_idx = None
            col_gubun = 1
            col_content = 2

            for idx in range(len(df)):
                row_vals = [re.sub(r'\s+', '', str(v)) for v in df.iloc[idx].values]
                if any("구분" in v for v in row_vals) and any("내용" in v for v in row_vals):
                    header_idx = idx
                    for c_idx, v in enumerate(row_vals):
                        if "구분" in v:
                            col_gubun = c_idx
                        elif "내용" in v and "의견" not in v:
                            col_content = c_idx
                    break

            visit_date = "일시 미상"
            file_issues = []
            file_requests = []
            current_section = None

            # 3. 표 본문 행 순회 (행 추가 및 병합 셀 완벽 대응)
            start_row = (header_idx + 1) if header_idx is not None else 0
            for idx in range(start_row, len(df)):
                row = df.iloc[idx]
                row_prefix_str = " ".join([re.sub(r'\s+', '', str(x)) for x in row.values[:col_content]])
                
                # 구분 키워드 인식 시 활성 섹션 전환
                if "일시" in row_prefix_str and not any(k in row_prefix_str for k in ["※", "【", "작성", "협의"]):
                    current_section = "일시"
                elif "현안문제" in row_prefix_str or "현안" in row_prefix_str:
                    current_section = "현안문제"
                elif "지원요청" in row_prefix_str or ("지원" in row_prefix_str and "요청" in row_prefix_str):
                    current_section = "지원요청사항"

                # 본문 내용 추출
                content_str = str(row[col_content]).strip() if col_content < len(row) else ""
                
                # 빈 셀, 기본 안내문구 건너뛰기
                if not content_str or content_str in ["내용 없음", "nan", "None"]:
                    continue

                if current_section == "일시":
                    visit_date = standardize_date(content_str)
                elif current_section == "현안문제":
                    file_issues.append(content_str)
                elif current_section == "지원요청사항":
                    file_requests.append(content_str)

            # 일정 목록 추가
            schedule_list.append({
                "level_sort": level_sort_val, 
                "학교급": school_level, 
                "학교명": raw_school_name, 
                "일시": visit_date, 
                "담당장학사": supervisor_name
            })
            
            # 현안문제 목록 추가 (추가된 모든 행 반영)
            for issue in file_issues:
                issue_list.append({
                    "level_sort": level_sort_val, 
                    "학교급": school_level, 
                    "학교명": raw_school_name, 
                    "현안문제": issue
                })
                
            # 지원요청사항 목록 추가 (추가된 모든 행 반영)
            for req in file_requests:
                request_list.append({
                    "level_sort": level_sort_val, 
                    "학교급": school_level, 
                    "학교명": raw_school_name, 
                    "지원요청사항": req
                })

            # 키워드 유목화 함수
            def classify_content(content, kind):
                if not content or content == "내용 없음":
                    return
                classified = False
                for category, keywords in categories.items():
                    if category == "기타(미분류)":
                        continue
                    if any(keyword in content for keyword in keywords):
                        categorized_list.append({
                            "level_sort": level_sort_val, 
                            "유목화 주제": category, 
                            "학교급": school_level, 
                            "학교명": raw_school_name, 
                            "구분": kind, 
                            "내용": content
                        })
                        classified = True
                        break 
                if not classified:
                    categorized_list.append({
                        "level_sort": level_sort_val, 
                        "유목화 주제": "기타(미분류)", 
                        "학교급": school_level, 
                        "학교명": raw_school_name, 
                        "구분": kind, 
                        "내용": content
                    })

            for issue in file_issues:
                classify_content(issue, "현안문제")
            for req in file_requests:
                classify_content(req, "지원요청사항")

        except Exception as e:
            st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

    # 데이터프레임 변환
    df_schedule = pd.DataFrame(schedule_list)
    df_issue = pd.DataFrame(issue_list)
    df_request = pd.DataFrame(request_list)
    df_categorized = pd.DataFrame(categorized_list)

    # 부서 요청 사항 생성
    dept_request_list = []
    for item in categorized_list:
        category = item["유목화 주제"]
        target_dept = dept_mapping.get(category, "관련 부서 확인 필요")
        dept_request_list.append({
            "level_sort": item["level_sort"],
            "유목화 주제": category,
            "담당 부서": target_dept,
            "학교급": item["학교급"],
            "학교명": item["학교명"],
            "구분": item["구분"],
            "요청 및 건의 내용": f"{item['내용']}\n\n[조치요청] 위 사항에 대한 구체적인 지원 방안 검토 요망"
        })
    df_dept_request = pd.DataFrame(dept_request_list)

    # 정렬 처리
    def sort_and_clean_default(df):
        if not df.empty:
            df = df.sort_values(by=['level_sort', '학교명'])
            df = df.drop(columns=['level_sort'])
        return df

    def sort_and_clean_categorized(df):
        if not df.empty:
            df = df.sort_values(by=['유목화 주제', '학교명'])
            df = df.drop(columns=['level_sort'])
        return df

    df_schedule = sort_and_clean_default(df_schedule)
    df_issue = sort_and_clean_default(df_issue)
    df_request = sort_and_clean_default(df_request)
    df_categorized = sort_and_clean_categorized(df_categorized)
    df_dept_request = sort_and_clean_categorized(df_dept_request)

    excel_sheets = {
        "1_방문일정": df_schedule,
        "2_학교현안문제": df_issue,
        "3_지원요청사항": df_request,
        "4_키워드유목화": df_categorized,
        "5_부서조치요청": df_dept_request
    }

    # 다운로드 버튼 및 탭 미리보기 구성
    st.divider()
    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.subheader("📁 데이터 추출 결과 및 통합 다운로드")
    with col_btn:
        st.download_button(
            label="📥 통합 엑셀 파일 다운로드", 
            data=to_excel_multi_sheet(excel_sheets), 
            file_name="지원장학_요청서_통합분석결과.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["방문 일정", "현안 문제", "지원 요청", "유목화 결과", "부서 조치 요청"])
    
    with tab1:
        st.dataframe(df_schedule, use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(df_issue, use_container_width=True, hide_index=True)
    with tab3:
        st.dataframe(df_request, use_container_width=True, hide_index=True)
    with tab4:
        st.dataframe(df_categorized, use_container_width=True, hide_index=True)
    with tab5:
        st.dataframe(df_dept_request, use_container_width=True, hide_index=True)
