import streamlit as st
import google.generativeai as genai
from pptx import Presentation
import json
import os
import tempfile
from io import BytesIO

# --- 1. 페이지 설정 및 디자인 ---
st.set_page_config(
    page_title="AI 예배 PPT 생성기",
    page_icon="🕊️",
    layout="centered"
)

# 깔끔한 UI를 위한 CSS
st.markdown("""
    <style>
    .main { padding-top: 2rem; }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        font-weight: bold;
    }
    .success-box {
        padding: 1rem;
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 0.5rem;
        color: #166534;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. 핵심 로직 (AI 및 PPT 처리) ---

def analyze_jubo_deep(image_file, key):
    """Gemini를 이용해 주보 정보 추출 및 성경 내용 생성"""
    genai.configure(api_key=key)
    # 요청하신 모델로 변경
    model = genai.GenerativeModel('gemini-3-flash-preview') 
    
    # Streamlit 업로드 파일을 임시 파일로 저장 (Gemini API 요구사항)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as tmp:
        tmp.write(image_file.getvalue())
        tmp_path = tmp.name

    try:
        sample_file = genai.upload_file(path=tmp_path)
        
        # 프롬프트: 주보 분석 + 성경 텍스트 생성 지시
        prompt = """
        이 주보 이미지에서 다음 정보를 찾아 JSON으로 출력해.
        값이 없으면 빈 문자열("")로 둬.
        
        1. sermon_title: 설교 제목
        2. preacher: 설교자 이름 (직분 포함, 예: 김철수 목사)
        3. prayer_person: 대표 기도자 이름
        4. bible_ref: 성경 본문 위치 (예: 요한복음 3:16)
        5. bible_text: 위 bible_ref에 해당하는 실제 성경 말씀 내용을 '개역개정' 버전으로 찾아서 전체 작성해줘. (인터넷 검색하지 말고 네가 아는 지식으로 정확하게)
        6. hymn_list: 찬송가 제목들을 순서대로 리스트에 담아줘. (예: ["찬송가 301장", "은혜"])
        """
        
        response = model.generate_content([sample_file, prompt])
        # JSON 포맷 정제
        text = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)
    except Exception as e:
        st.error(f"분석 중 오류가 발생했습니다: {e}")
        return None
    finally:
        # 임시 파일 삭제
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def fill_ppt_text(template_file, data):
    """PPT 텍스트 교체 (디자인 서식 유지)"""
    prs = Presentation(template_file)
    
    # 템플릿과 매칭될 데이터 사전
    replacements = {
        "{{설교제목}}": data.get('sermon_title', ''),
        "{{설교자}}": data.get('preacher', ''),
        "{{기도자}}": data.get('prayer_person', ''),
        "{{성경본문}}": data.get('bible_ref', ''),
        "{{말씀내용}}": data.get('bible_text', '')
    }
    
    # 찬송가 리스트 처리 ({{찬송1}}, {{찬송2}}...)
    hymns = data.get('hymn_list', [])
    for i, hymn in enumerate(hymns):
        replacements[f"{{{{찬송{i+1}}}}}"] = hymn
    
    # 모든 슬라이드 -> 모든 도형 -> 모든 문단 -> 모든 런(Run) 순회
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for k, v in replacements.items():
                        if k in run.text:
                            # 값이 없으면 빈칸, 있으면 문자열로 변환하여 교체
                            safe_value = str(v) if v is not None else ""
                            run.text = run.text.replace(k, safe_value)

    # 결과를 바이너리 스트림으로 저장
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- 3. 화면 UI 구성 ---

st.title("🕊️ AI 예배 PPT 생성기")
st.markdown("주보 사진만 올리면, 템플릿의 빈칸을 AI가 자동으로 채워줍니다.")

# 사이드바: 설정 및 도움말
with st.sidebar:
    st.header("⚙️ 설정")
    api_key = st.text_input("Google API Key", type="password", placeholder="API 키를 입력하세요")
    st.caption("[API 키 발급받기](https://aistudio.google.com/app/apikey)")
    
    st.divider()
    
    with st.expander("❓ 템플릿 만드는 법 (필독)"):
        st.markdown("""
        PPT 템플릿의 텍스트 상자에 아래 **단어**를 적어두세요.
        AI가 이 단어를 찾아 내용으로 바꿔치기합니다.
        
        - `{{설교제목}}`
        - `{{설교자}}`
        - `{{기도자}}`
        - `{{성경본문}}` (예: 요 3:16)
        - `{{말씀내용}}` (성경 구절이 자동으로 들어감)
        - `{{찬송1}}`, `{{찬송2}}`...
        """)
        st.info("중괄호 {{ }}를 꼭 두 번 겹쳐서 써야 합니다!")

# 메인 기능 영역
if not api_key:
    st.warning("👈 왼쪽 사이드바에 Google API 키를 먼저 입력해주세요.")
else:
    # STEP 1: 파일 업로드
    st.subheader("1. 파일 업로드")
    col1, col2 = st.columns(2)
    with col1:
        jubo_img = st.file_uploader("주보 이미지 (사진)", type=['png', 'jpg', 'jpeg'])
    with col2:
        template_pptx = st.file_uploader("PPT 템플릿 파일", type=['pptx'])

    # STEP 2: AI 분석 실행
    if jubo_img and template_pptx:
        st.divider()
        # 버튼을 누르면 분석 시작
        if st.button("주보 분석 시작 ✨", type="primary"):
            with st.spinner("주보를 읽고 성경 말씀을 찾는 중입니다..."):
                result = analyze_jubo_deep(jubo_img, api_key)
                if result:
                    # 결과를 세션 상태에 저장 (새로고침 방지)
                    st.session_state['ppt_data'] = result
                    st.rerun()

    # STEP 3: 결과 확인 및 다운로드
    if 'ppt_data' in st.session_state:
        st.divider()
        st.subheader("2. 내용 확인 및 수정")
        
        st.markdown('<div class="success-box">✅ AI 분석 완료! 내용을 확인하세요.</div>', unsafe_allow_html=True)
        
        # 데이터 가져오기
        d = st.session_state['ppt_data']
        
        # 수정 가능한 폼(Form) 생성
        with st.form("check_form"):
            c1, c2 = st.columns(2)
            with c1:
                new_title = st.text_input("설교 제목", value=d.get('sermon_title', ''))
                new_preacher = st.text_input("설교자", value=d.get('preacher', ''))
            with c2:
                new_prayer = st.text_input("기도자", value=d.get('prayer_person', ''))
                new_ref = st.text_input("성경 본문", value=d.get('bible_ref', ''))
            
            new_text = st.text_area("성경 말씀 내용 (AI 자동 생성)", value=d.get('bible_text', ''), height=150)
            
            # 리스트를 문자열로 변환하여 표시
            hymn_str = st.text_input("찬송가 순서 (쉼표로 구분)", value=", ".join(d.get('hymn_list', [])))
            
            # 생성 버튼
            submitted = st.form_submit_button("이 내용으로 PPT 만들기 🎁", type="primary")
            
            if submitted:
                # 최종 데이터 정리
                final_data = {
                    "sermon_title": new_title,
                    "preacher": new_preacher,
                    "prayer_person": new_prayer,
                    "bible_ref": new_ref,
                    "bible_text": new_text,
                    "hymn_list": [h.strip() for h in hymn_str.split(',')]
                }
                
                # PPT 생성 함수 호출
                final_ppt = fill_ppt_text(template_pptx, final_data)
                
                # 다운로드 버튼 표시 (폼 밖으로 나가기 위해 세션 사용 권장하지만, 여기선 바로 표시)
                st.divider()
                st.balloons()
                st.success("작업이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
                
                st.download_button(
                    label="📥 완성된 PPT 다운로드",
                    data=final_ppt,
                    file_name=f"{new_title}_예배.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
