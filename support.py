import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Alignment

# 다중 시트 엑셀 다운로드를 위한 변환 함수
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

# 일시 포맷 표준화 ('0000년 00월 00일 00:00')
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

# 상단 헤더 영역에서 학교명, 학교급, 담당 장학사 추출
def extract_header_info(df, filename=""):
    raw_school_name = "학교명 미상"
    supervisor_name = "장학사 미상"
    school_level = ""
    level_sort_val = 3

    for idx in range(min(12, len(df))):
        val = str(df.iloc[idx, 0]).strip()
        if not val or val.startswith("【") or val.startswith("※") or val.startswith("작성") or val.startswith("("):
            continue
        if val == "순":
            break
            
        if "장학사" in val:
            sup_match = re.search(r'(?:학교담당장학사|담당장학사|장학사)\s*[:：]?\s*([가-힣]{2,4})', val)
            if sup_match and supervisor_name == "장학사 미상":
                supervisor_name = sup_match.group(1).strip()
        
        if ("학교" in val or "중" in val or "고" in val) and "장학사" not in val:
            if raw_school_name == "학교명 미상":
                raw_school_name = val

    if raw_school_name == "학교명 미상" and filename:
        parts = filename.replace(".xlsx", "").replace(".csv", "").split("_")
        if len(parts) > 1:
            raw_school_name = parts[1]

    if "중학교" in raw_school_name or filename.startswith("중_"):
        school_level = "중"
        level_sort_val = 1
    elif "고등학교" in raw_school_name or filename.startswith("고_"):
        school_level = "고"
        level_sort_val = 2

    return raw_school_name, school_level, level_sort_val, supervisor_name

# 9대 핵심 연계 영역 및 부서 업무 매핑 정보
DOMAIN_AREAS = [
    {
        "area_no": 1,
        "area": "교육과정 및 수업·평가 (고교학점제 포함)",
        "main_dept": "중등교육지원과",
        "sub_dept": "학교통합지원과, 평생교육건강과",
        "work_desc": "중등 교육과정, 고교학점제 운영 지원, 성적·수행평가 관리, 최성보(최소성취수준보장지도), 학생부 기재요령, 수능/학평 관리 | (연계) 교과서 배부 지원(학교통합), 학원 교습시간 및 시험지 유출 지도점검(평생건강)",
        "keywords": [
            "고교학점제", "교육과정", "최소성취수준", "최성보", "학생부", "세특", "생활기록부",
            "성적", "학업성적", "과정중심", "IB", "수업나눔", "수능", "모의평가", "학평",
            "고입", "진로", "진학", "교과서배부", "수업", "평가", "운동부", "도서관"
        ]
    },
    {
        "area_no": 2,
        "area": "특수교육 및 느린학습자 지원",
        "main_dept": "중등교육지원과 (총괄지원센터: 초등)",
        "sub_dept": "행정지원과, 학생맞춤협력과, 학교통합지원과",
        "work_desc": "중등 특수교육 운영 및 장학(중등), 특수교육대상자 진단·배치 및 순회교육(특수교육지원센터) | (연계) 특수학급 신·증설 계획(행정지원), 경계선 지능·난독 진단(학습진단성장센터/학생맞춤), 특수실무사 관리(학교통합)",
        "keywords": [
            "특수교육", "특수학급", "특수실무사", "느린학습자", "경계선지능", "난독", "기초학력", "학습진단", "학습지원튜터"
        ]
    },
    {
        "area_no": 3,
        "area": "생활지도 및 심리·정서 지원 (학생맞춤통합지원)",
        "main_dept": "학생맞춤협력과, 학교생활교육과",
        "sub_dept": "행정지원과",
        "work_desc": "위기학생 상담, Wee센터 운영, 학생맞춤통합지원(학맞통), 복합위기 학생 전문기관 연계(학생맞춤), 학생 생활교육 규정(학칙), 학교폭력 사안 처리(전담조사관·학폭위), 마음건강 증진(자살예방) | (연계) 미인정결석 학생 유관기관 협조(행정지원)",
        "keywords": [
            "학맞통", "학생맞춤통합", "위기학생", "Wee", "위클래스", "마음건강", "자살", "자해",
            "생명존중", "정서", "심리", "상담", "학교폭력", "학폭", "전담조사관", "생활지도",
            "생활규정", "학칙", "출결", "결석", "미인정결석", "다문화", "교육복지", "햇살교실", "사회정서"
        ]
    },
    {
        "area_no": 4,
        "area": "교육활동 보호 및 대외(민원) 대응",
        "main_dept": "학교생활교육과",
        "sub_dept": "중등교육지원과, 행정지원과",
        "work_desc": "교육활동보호지원센터(SEM119), 지역교권보호위원회 심의·운영, 교원 대상 아동학대 신고 조사·대응, 교권침해 법률·상담 자문 | (연계) 악성 민원에 따른 교원 복무 및 장학 지원(중등)",
        "keywords": [
            "교권", "교육활동보호", "SEM119", "교권보호위원회", "악성민원", "아동학대신고", "법률지원", "도난", "분실"
        ]
    },
    {
        "area_no": 5,
        "area": "교원 인사 및 학교 인력 확충",
        "main_dept": "중등교육지원과, 학교통합지원과",
        "sub_dept": "학생맞춤협력과",
        "work_desc": "중등 정규 교원 정원·배치, 전보 및 초빙, 순회교사제 운영(중등), 학교 기간제교원·시간강사 인력풀 및 채용 지원, 교육공무직원(교무행정사·조리실무사 등) 정원·전보 관리(학교통합) | (연계) 전문상담교사 배치 관리(학생맞춤)",
        "keywords": [
            "교원인사", "정원", "전보", "순회교사", "기간제", "시간강사", "인력풀", "채용",
            "교육공무직", "조리실무사", "교무행정", "호봉", "인력배정", "상담교사증원", "인력지원"
        ]
    },
    {
        "area_no": 6,
        "area": "디지털 교육환경 및 정보화 지원",
        "main_dept": "학교통합지원과",
        "sub_dept": "중등교육지원과, 행정지원과",
        "work_desc": "디지털 튜터 배치·운영, 디벗(스마트기기) 회수·재배부 관리, 학교 정보화 인프라 및 네트워크 테크센터 운영 지원(학교통합) | (연계) AI 중점학교 운영 및 에듀테크 연수(중등), 정보보안 수준진단 및 나이스 시스템 연계(행정지원)",
        "keywords": [
            "디벗", "스마트기기", "디지털튜터", "테크센터", "정보화", "인프라", "네트워크",
            "전자칠판", "AI스튜디오", "스튜디오", "정보보호", "개인정보", "사이버침해", "에듀테크", "코딩", "디지털"
        ]
    },
    {
        "area_no": 7,
        "area": "학교시설 및 교육환경 개선",
        "main_dept": "학교시설지원과",
        "sub_dept": "평생교육건강과, 재정지원과",
        "work_desc": "학교 누수(옥상·연결통로 방수), 냉난방기 교체, 교내 방송설비 개선, 내진보강, 특별실·교무실 공간재구조화, 소음차단 시설, BTL 관리점검 | (연계) 급식실 환경개선 및 위탁급식 컨설팅, 보건실 환경, 통학로 교육환경보호구역(평생건강), 시설공사 및 물품 계약·집행(재정지원)",
        "keywords": [
            "누수", "방수", "지붕", "냉난방", "방송설비", "내진보강", "석면", "도색",
            "공간재구조화", "공간혁신", "특별실", "시설안전", "정밀안전", "옹벽", "항공소음",
            "BTL", "공사", "노후", "수리", "급식실", "위탁급식", "운반급식", "급식", "보건실", "통학로"
        ]
    },
    {
        "area_no": 8,
        "area": "학생배치 및 학교규모 적정화",
        "main_dept": "행정지원과 (학생배치팀·목동재건축팀)",
        "sub_dept": "중등교육지원과",
        "work_desc": "중장기 학생수용계획, 학급편제 및 학급당 학생수(과밀/과소) 조정, 중학교 입학 및 전·편입학 배정, 신설학교 개교에 따른 학생 재배치, 목동재건축지구 학교 설립·배치 계획 | (연계) 학급 수 변동에 따른 교원 정원 조정(중등)",
        "keywords": [
            "학생배치", "학생수용", "학급편제", "과밀학급", "과밀", "배정", "신입생배정",
            "전편입", "신설학교", "학교설립", "학교폐지", "재건축"
        ]
    },
    {
        "area_no": 9,
        "area": "학교재정 및 행정 업무 경감",
        "main_dept": "재정지원과, 학교통합지원과",
        "sub_dept": "중등교육지원과, 행정지원과, 평생교육건강과",
        "work_desc": "학교회계 예산 편성·교부(목적사업비), 계약 체결 및 채권 관리(재정지원), 교과서 일괄 배부 지원 업체 관리감독, 학교 행정업무 경감 지원(학교통합) | (연계) 항공소음 피해학교 냉방비 지원(평생건강/시설지원), 에듀파인 공문·시스템 개선(행정지원)",
        "keywords": [
            "학교회계", "목적사업비", "교육비특별회계", "예산", "계약", "품의", "결제",
            "용역", "공유재산", "소송", "채권", "유아학비", "급여", "지출", "업무경감", "에듀파인"
        ]
    }
]

# 본문 내용 기반 영역 유목화 함수
def match_domain_area(content):
    if not content or content in ["내용 없음", "nan", "None"]:
        return None
    for dom in DOMAIN_AREAS:
        if any(kw in content for kw in dom["keywords"]):
            return dom
    return {
        "area_no": 99,
        "area": "기타(미분류)",
        "main_dept": "관련 부서 확인 필요",
        "sub_dept": "확인 필요",
        "work_desc": "세부 업무 내용 확인 후 주관/협조 부서 지정 필요"
    }

# Streamlit 페이지 설정
st.set_page_config(page_title="지원장학 요청서 자동 분석기", layout="wide")
st.title("📊 지원장학 요청서 자동 분석 및 부서 연계 웹앱")
st.markdown("지원장학 요청서를 업로드하면 지정된 셀에서 데이터를 추출하고, **교육지원청 9대 연계 영역 및 부서별 업무 체계**에 따라 자동 분류·정리합니다.")

uploaded_files = st.file_uploader("장학 요청서 파일(Excel 또는 CSV)을 업로드하세요.", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("분석 시작") and uploaded_files:
    schedule_list = []
    issue_list = []
    request_list = []
    categorized_list = []
    dept_request_list = []

    st.success(f"총 {len(uploaded_files)}개의 파일을 분석합니다...")
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, header=None, dtype=str).fillna("")
            else:
                df = pd.read_excel(file, header=None, dtype=str).fillna("")

            # 1. 학교 기본정보 추출
            raw_school_name, school_level, level_sort_val, supervisor_name = extract_header_info(df, file.name)

            # 2. 표 헤더('구분', '내용') 위치 동적 감지
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

            # 3. 본문 행 순회 (행 추가 및 셀 병합 완벽 대응)
            start_row = (header_idx + 1) if header_idx is not None else 0
            for idx in range(start_row, len(df)):
                row = df.iloc[idx]
                row_prefix_str = " ".join([re.sub(r'\s+', '', str(x)) for x in row.values[:col_content]])
                
                if "일시" in row_prefix_str and not any(k in row_prefix_str for k in ["※", "【", "작성", "협의"]):
                    current_section = "일시"
                elif "현안문제" in row_prefix_str or "현안" in row_prefix_str:
                    current_section = "현안문제"
                elif "지원요청" in row_prefix_str or ("지원" in row_prefix_str and "요청" in row_prefix_str):
                    current_section = "지원요청사항"

                content_str = str(row[col_content]).strip() if col_content < len(row) else ""
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
            
            # 현안문제 목록 추가
            for issue in file_issues:
                issue_list.append({
                    "level_sort": level_sort_val, 
                    "학교급": school_level, 
                    "학교명": raw_school_name, 
                    "현안문제": issue
                })
                
            # 지원요청사항 목록 추가
            for req in file_requests:
                request_list.append({
                    "level_sort": level_sort_val, 
                    "학교급": school_level, 
                    "학교명": raw_school_name, 
                    "지원요청사항": req
                })

            # 키워드 유목화 및 부서 연계 매핑
            def process_classification(content, kind):
                matched = match_domain_area(content)
                if not matched:
                    return
                
                # 키워드 유목화 데이터
                categorized_list.append({
                    "level_sort": level_sort_val,
                    "area_sort": matched["area_no"],
                    "유목화 영역": matched["area"],
                    "주관 담당과": matched["main_dept"],
                    "협조·연계 담당과": matched["sub_dept"],
                    "학교급": school_level,
                    "학교명": raw_school_name,
                    "구분": kind,
                    "내용": content
                })
                
                # 부서 조치 요청 데이터
                dept_request_list.append({
                    "level_sort": level_sort_val,
                    "area_sort": matched["area_no"],
                    "유목화 영역": matched["area"],
                    "주관 담당과": matched["main_dept"],
                    "협조·연계 담당과": matched["sub_dept"],
                    "학교급": school_level,
                    "학교명": raw_school_name,
                    "구분": kind,
                    "요청 및 건의 내용": f"{content}\n\n[조치요청] 위 사항에 대한 구체적인 지원 방안 검토 요망",
                    "주요 연계 업무 안내": matched["work_desc"]
                })

            for issue in file_issues:
                process_classification(issue, "현안문제")
            for req in file_requests:
                process_classification(req, "지원요청사항")

        except Exception as e:
            st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

    # DataFrame 변환
    df_schedule = pd.DataFrame(schedule_list)
    df_issue = pd.DataFrame(issue_list)
    df_request = pd.DataFrame(request_list)
    df_categorized = pd.DataFrame(categorized_list)
    df_dept_request = pd.DataFrame(dept_request_list)

    # 정렬 함수
    def sort_and_clean_default(df):
        if not df.empty:
            df = df.sort_values(by=['level_sort', '학교명'])
            df = df.drop(columns=['level_sort'])
        return df

    def sort_and_clean_categorized(df):
        if not df.empty:
            df = df.sort_values(by=['area_sort', 'level_sort', '학교명'])
            df = df.drop(columns=['area_sort', 'level_sort'])
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

    # 다운로드 및 화면 탭 출력
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
