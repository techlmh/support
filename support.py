import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Alignment

# 다중 시트 엑셀 다운로드를 위한 변환 함수 (줄바꿈 서식 적용 추가)
def to_excel_multi_sheet(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            
            # 생성된 시트 객체 가져오기
            worksheet = writer.sheets[sheet_name]
            
            # 모든 셀에 대해 '텍스트 줄바꿈(wrap_text)' 서식 적용
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                    
            # 열 너비 자동 조절 (내용이 잘 보이도록)
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter # 열 알파벳 (A, B, C...)
                for cell in col:
                    try:
                        # 줄바꿈이 있는 경우 가장 긴 줄을 기준으로 너비 계산
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            if len(line) > max_length:
                                max_length = len(line)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 80) # 최대 너비 80으로 제한
                worksheet.column_dimensions[column].width = adjusted_width

    processed_data = output.getvalue()
    return processed_data

# 일시 양식을 '0000년 00월 00일 00:00 (24시간제)'로 통일하는 함수
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

# 페이지 기본 설정
st.set_page_config(page_title="지원장학 요청서 자동 분석기", layout="wide")
st.title("📊 지원장학 요청서 자동 분석 및 유목화 웹앱")
st.markdown("지원장학 요청서를 업로드하면 지정된 셀에서 데이터를 추출하고, 조건에 맞게 정렬하여 엑셀 파일로 추출합니다.")

# 파일 업로더
uploaded_files = st.file_uploader("장학 요청서 파일(Excel 또는 CSV)을 업로드하세요.", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("분석 시작") and uploaded_files:
    schedule_list = []
    issue_list = []
    request_list = []
    categorized_list = []
    
    categories = {
        "시설 및 환경 개선": ["방송", "공사", "누수", "노후", "수리", "교체", "안전", "공간", "장비"],
        "예산 및 행정 지원": ["예산", "지원금", "품의", "결제", "계약", "행정", "인력", "채용", "강사"],
        "교육과정 및 학사 운영": ["교과", "학점제", "평가", "성적", "교육과정", "자유학기", "수업", "디지털", "코딩"],
        "생활지도 및 학생 지원": ["폭력", "학폭", "상담", "정서", "위기", "징계", "출결", "다문화"],
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

            # 1. 학교명 및 학교급 추출
            raw_school_name = "학교명 미상"
            if len(df) > 3 and df.iloc[3, 0]:
                raw_school_name = str(df.iloc[3, 0]).strip()

            school_level = ""
            level_sort_val = 3
            if "중학교" in raw_school_name:
                school_level = "중"
                level_sort_val = 1
            elif "고등학교" in raw_school_name:
                school_level = "고"
                level_sort_val = 2

            # 2. 장학사 이름 추출
            supervisor_name = "장학사 미상"
            if len(df) > 4 and df.iloc[4, 0]:
                a5_text = str(df.iloc[4, 0]).strip()
                if len(a5_text) >= 3:
                    supervisor_name = a5_text[-3:]
                else:
                    supervisor_name = a5_text

            # 3. 일시 및 표 내용 추출 (내용 안의 줄바꿈 유지)
            visit_date = "일시 미상"
            issue_content = ""
            support_content = ""

            for index, row in df.iterrows():
                row_list = list(row.values)
                for col_idx, cell_value in enumerate(row_list):
                    if "일시" in str(cell_value):
                        visit_date = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else visit_date
                    elif "현안문제" in str(cell_value):
                        issue_content = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else issue_content
                    elif "지원 요청 사항" in str(cell_value):
                        support_content = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else support_content

            visit_date = standardize_date(visit_date)
            # strip()으로 앞뒤 공백만 제거하고, 내부 줄바꿈(\n)은 그대로 보존
            issue_content = issue_content.strip()
            support_content = support_content.strip()

            # 데이터 추가 (학교급과 학교명을 별도의 키로 분리)
            schedule_list.append({"level_sort": level_sort_val, "학교급": school_level, "학교명": raw_school_name, "일시": visit_date, "담당장학사": supervisor_name})
            
            if issue_content:
                issue_list.append({"level_sort": level_sort_val, "학교급": school_level, "학교명": raw_school_name, "현안문제": issue_content})
            if support_content:
                request_list.append({"level_sort": level_sort_val, "학교급": school_level, "학교명": raw_school_name, "지원요청사항": support_content})

            def classify_content(content, kind, school_lvl, school_nm, sort_val):
                if not content or content == "내용 없음":
                    return
                classified = False
                for category, keywords in categories.items():
                    if category == "기타(미분류)":
                        continue
                    if any(keyword in content for keyword in keywords):
                        categorized_list.append({"level_sort": sort_val, "유목화 주제": category, "학교급": school_lvl, "학교명": school_nm, "구분": kind, "내용": content})
                        classified = True
                        break 
                if not classified:
                    categorized_list.append({"level_sort": sort_val, "유목화 주제": "기타(미분류)", "학교급": school_lvl, "학교명": school_nm, "구분": kind, "내용": content})

            classify_content(issue_content, "현안문제", school_level, raw_school_name, level_sort_val)
            classify_content(support_content, "지원요청사항", school_level, raw_school_name, level_sort_val)

        except Exception as e:
            st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

    # 리스트를 Pandas DataFrame으로 변환
    df_schedule = pd.DataFrame(schedule_list)
    df_issue = pd.DataFrame(issue_list)
    df_request = pd.DataFrame(request_list)
    df_categorized = pd.DataFrame(categorized_list)

    # 부서 요청 사항 데이터프레임 생성 (요청 내용에 줄바꿈 추가)
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
            # 원본 내용 아래에 줄바꿈(\n)을 넣어 조치 요청 멘트 추가
            "요청 및 건의 내용": f"{item['내용']}\n\n[조치요청] 위 사항에 대한 구체적인 지원 방안 검토 요망"
        })
    df_dept_request = pd.DataFrame(dept_request_list)

    # --- 정렬 로직 분리 ---
    # 1. 일반 시트 정렬 (학교급 -> 학교명 순서)
    def sort_and_clean_default(df):
        if not df.empty:
            df = df.sort_values(by=['level_sort', '학교명'])
            df = df.drop(columns=['level_sort'])
        return df

    # 2. 유목화 시트 정렬 (유목화 주제 -> 학교명 순서)
    def sort_and_clean_categorized(df):
        if not df.empty:
            df = df.sort_values(by=['유목화 주제', '학교명'])
            df = df.drop(columns=['level_sort'])
        return df

    df_schedule = sort_and_clean_default(df_schedule)
    df_issue = sort_and_clean_default(df_issue)
    df_request = sort_and_clean_default(df_request)
    
    # 키워드 유목화 및 부서 조치 요청은 "유목화 주제 -> 학교명" 기준으로 정렬
    df_categorized = sort_and_clean_categorized(df_categorized)
    df_dept_request = sort_and_clean_categorized(df_dept_request)

    # 엑셀 파일 생성을 위한 딕셔너리 매핑
    excel_sheets = {
        "1_방문일정": df_schedule,
        "2_학교현안문제": df_issue,
        "3_지원요청사항": df_request,
        "4_키워드유목화": df_categorized,
        "5_부서조치요청": df_dept_request
    }

    # --- 화면 출력 및 통합 엑셀 다운로드 ---
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

    # 탭을 사용하여 화면을 깔끔하게 구성 (데이터 미리보기)
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
