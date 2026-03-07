import streamlit as st
import streamlit.components.v1 as components
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from bs4 import BeautifulSoup
import pandas as pd
import base64
import time
import re
import tempfile
import shutil

TEXTS = {
    "zh": {
        "title": "東華大學 GPA 計算機",
        "desc": "本工具自動擷取校務系統資料，僅供個人參考不能作為成績證明。\n\n此外每所學校的成績計算方式會有差異，仍須以要申請的校系公告為準。\n\n問題回報表單：https://forms.gle/tidiWbmgdaQY3rCm6",
        "login_header": "系統登入",
        "user_id": "學號",
        "user_pw": "密碼",
        "captcha_btn": "取得驗證碼",
        "captcha_input": "輸入上方驗證碼",
        "login_btn": "登入",
        "logout_btn": "登出",
        "loading_sys": "正在連線至系統...",
        "auth_ing": "正在進行身分驗證...",
        "fetch_ing": "驗證成功，正在讀取成績分頁...",
        "parsing": "正在檢索成績清單...",
        "success": "資料分析完成。",
        "error_timeout": "連線逾時或系統無回應，無法讀取驗證碼。",
        "error_login": "登入失敗，請確認登入資訊是否正確。（錯誤次數：{count}/3）",
        "error_process": "分析過程異常中斷，請稍後再試。（錯誤碼：{err}）",
        "summary_header": "歷年總成績 (Overall)",
        "last4_header": "最後四學期成績 (Last 4 Semesters)",
        "last60_header": "最後 60 學分成績 (Last 60 Credits)",
        "overall_gpa": "GPA",
        "detail_tab": "查看詳細成績明細清單",
        "col_term": "學期",
        "col_name": "科目名稱",
        "col_req": "必/選修",
        "col_credit": "學分",
        "col_grade": "取得等第",
        "conv_header": "查看 GPA 換算對照表",
        "conv_grade": "等第",
        "conv_45_col": "4.5 制",
        "conv_40_col": "4.0 制",
        "conv_43_col": "4.3 制",
        "privacy": "隱私說明：本工具僅於工作階段處理數據，關閉分頁或登出後系統會自動刪除所有暫存與紀錄。",
        "lockout": "登入失敗次數過多，請等候 {time} 秒後再試。"
    },
    "en": {
        "title": "NDHU GPA Calculator",
        "desc": "This tool automatically retrieves data from the university system for personal reference only and cannot be used as an official transcript.\n\nGrading scales may vary by institution; please refer to the official announcements of the target university. \n\nIssue reporting form: https://forms.gle/tidiWbmgdaQY3rCm6",
        "login_header": "System Login",
        "user_id": "Student ID",
        "user_pw": "Password",
        "captcha_btn": "Get Captcha",
        "captcha_input": "Enter Captcha",
        "login_btn": "Login",
        "logout_btn": "Logout",
        "loading_sys": "Connecting to system...",
        "auth_ing": "Authenticating...",
        "fetch_ing": "Login success, accessing grades...",
        "parsing": "Retrieving grade list...",
        "success": "Analysis completed.",
        "error_timeout": "Connection timeout, failed to load captcha.",
        "error_login": "Login failed. (Attempts: {count}/3)",
        "error_process": "Analysis interrupted. (Err: {err})",
        "summary_header": "Overall Performance",
        "last4_header": "Last 4 Semesters",
        "last60_header": "Last 60 Credits",
        "overall_gpa": "GPA",
        "detail_tab": "View Detailed Course List",
        "col_term": "Term",
        "col_name": "Course Name",
        "col_req": "Req./Elec.",
        "col_credit": "Credits",
        "col_grade": "Grade",
        "conv_header": "GPA Conversion Table",
        "conv_grade": "Grade",
        "conv_45_col": "4.5 Scale",
        "conv_40_col": "4.0 Scale",
        "conv_43_col": "4.3 Scale",
        "privacy": "Privacy: This tool processes data only during the session. All temporary data will be automatically deleted after closing the page or logging out.",
        "lockout": "Too many failed attempts. Try again in {time} seconds."
    }
}

LOGIN_URL = 'https://sys.ndhu.edu.tw/CTE/Ed_StudP_WebSite/Login.aspx'

def compute_gpa_analytics(parsed_rows):
    convert_45 = {
        'A+': 4.5, 'A': 4.0, 'A-': 3.7, 'B+': 3.3, 'B': 3.0, 'B-': 2.7,
        'C+': 2.5, 'C': 2.3, 'C-': 2.0, 'D': 1.0, 'E': 0.0
    }
    convert_40 = {
        'A+': 4.0, 'A': 4.0, 'A-': 3.67, 'B+': 3.33, 'B': 3.0, 'B-': 2.67,
        'C+': 2.33, 'C': 2.0, 'C-': 1.67, 'D': 1.33, 'E': 0.0
    }
    convert_43 = {
        'A+': 4.3, 'A': 4.0, 'A-': 3.7, 'B+': 3.3, 'B': 3.0, 'B-': 2.7,
        'C+': 2.3, 'C': 2.0, 'C-': 1.7, 'D': 1.3, 'E': 0.0
    }

    valid_courses = []
    
    for row in parsed_rows:
        n = row.get('name', '')
        c = row.get('credit', 0.0)
        g = row.get('grade', '')
        y = row.get('year', '0')
        s = row.get('seme', '0')
        req = row.get('req_elec', '').strip()

        if '操行' in str(n) or 'Conduct' in str(n) or (float(c) <= 0):
            continue
            
        grade_val = str(g).strip().upper()
        if grade_val == 'W':
            continue
            
        if grade_val not in convert_40:
            if re.search(r'[甲乙丙丁優良可|()]+', grade_val):
                continue
            continue

        try: term_y = int(y)
        except: term_y = 0
        try: term_s = int(s)
        except: term_s = 0

        valid_courses.append({
            'year': term_y,
            'seme': term_s,
            'name': str(n),
            'req_elec': req,
            'credit': float(c),
            'grade': grade_val,
            'gp45': convert_45[grade_val],
            'gp40': convert_40[grade_val],
            'gp43': convert_43[grade_val]
        })

    def calc_stats(courses):
        c_total = sum(x['credit'] for x in courses)
        g45 = round(sum(x['credit']*x['gp45'] for x in courses) / c_total, 2) if c_total else 0.0
        g40 = round(sum(x['credit']*x['gp40'] for x in courses) / c_total, 2) if c_total else 0.0
        g43 = round(sum(x['credit']*x['gp43'] for x in courses) / c_total, 2) if c_total else 0.0
        return g45, g40, g43, c_total

    overall_stats = calc_stats(valid_courses)

    unique_terms = sorted(list(set((x['year'], x['seme']) for x in valid_courses)), reverse=True)
    last_4_terms = unique_terms[:4]
    l4_courses = [x for x in valid_courses if (x['year'], x['seme']) in last_4_terms]
    last4_stats = calc_stats(l4_courses)

    sorted_courses = sorted(valid_courses, key=lambda x: (x['year'], x['seme']))
    l60_courses = []
    acc_c = 0.0
    for x in sorted_courses[::-1]:
        l60_courses.append(x)
        acc_c += x['credit']
        if acc_c >= 60: break
    last60_stats = calc_stats(l60_courses)

    return overall_stats, last4_stats, last60_stats, valid_courses

def init_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions") 
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    options.add_argument("--incognito")
    options.add_argument("--disable-application-cache")
    options.add_argument("--disk-cache-size=0")
    options.add_argument("--log-level=3") 
    
    temp_profile_dir = tempfile.mkdtemp()
    st.session_state.temp_profile_dir = temp_profile_dir
    options.add_argument(f"--user-data-dir={temp_profile_dir}")
    
    driver_path = shutil.which("chromedriver") 
    if driver_path:
        service = Service(driver_path)
    else:
        service = Service("/usr/bin/chromedriver")
        
    try:
        return webdriver.Chrome(service=service, options=options)
    except Exception:
        return webdriver.Chrome(options=options)

def cleanup_driver():
    if 'driver_instance' in st.session_state and st.session_state.driver_instance is not None:
        try:
            st.session_state.driver_instance.quit()
        except Exception:
            pass
        finally:
            st.session_state.driver_instance = None
            
    if 'temp_profile_dir' in st.session_state and st.session_state.temp_profile_dir is not None:
        try:
            shutil.rmtree(st.session_state.temp_profile_dir, ignore_errors=True)
        except Exception:
            pass
        finally:
            st.session_state.temp_profile_dir = None

st.set_page_config(page_title="NDHU GPA Calculator", layout="wide")

st.markdown("""
    <style>
    div.stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
        background-color: #e53935 !important;
        color: white !important;
        border-color: #e53935 !important;
    }
    div.stButton > button[kind="primary"]:hover, div[data-testid="stFormSubmitButton"] > button[kind="primary"]:hover {
        background-color: #c62828 !important;
        border-color: #c62828 !important;
    }
    
    [data-testid="stMetricLabel"] p {
        font-size: 1.15rem !important;
        font-weight: 600 !important;
    }

    .custom-fixed-header {
        position: fixed !important;
        top: 2.8rem !important;
        left: 0 !important;
        right: 0 !important;
        z-index: 9999 !important;
        
        background-color: color-mix(in srgb, var(--background-color) 80%, transparent) !important;
        backdrop-filter: blur(8px) !important;
        -webkit-backdrop-filter: blur(8px) !important;
        
        border-bottom: 1px solid color-mix(in srgb, var(--text-color) 15%, transparent) !important;
        padding: 1rem 3rem !important; 
        transition: padding 0.3s ease !important; 
    }

    .desktop-spacer { height: 2.2rem; }
    .header-spacer { height: 1.5rem; }

    @media (max-width: 1024px) {
        .custom-fixed-header { padding: 1rem 2rem !important; }
    }

    @media (max-width: 768px) {
        .custom-fixed-header {
            top: 2rem !important;
            gap: 0.2rem !important;
        }
        .custom-fixed-header h1 {
            font-size: 1.6rem !important; 
            padding-top: 0rem !important;
            padding-bottom: 0 !important;
            margin-bottom: 0 !important;
        }
        .desktop-spacer { display: none; }
        .header-spacer { height: 0rem; }
    }
    </style>
""", unsafe_allow_html=True)

if 'lang' not in st.session_state:
    st.session_state.lang = 'zh'
if 'driver_instance' not in st.session_state:
    st.session_state.driver_instance = None
if 'captcha_bytes' not in st.session_state:
    st.session_state.captcha_bytes = None
if 'fail_count' not in st.session_state:
    st.session_state.fail_count = 0
if 'lockout_until' not in st.session_state:
    st.session_state.lockout_until = 0

T = TEXTS[st.session_state.lang]
has_data = 'final_parsed_rows' in st.session_state

col_title, col_logout = st.columns([5, 1])

with col_title:
    st.markdown('<div id="main-header-marker" style="display: none;"></div>', unsafe_allow_html=True)
    st.title(T["title"])
    
with col_logout:
    if has_data:
        st.markdown('<div class="desktop-spacer"></div>', unsafe_allow_html=True)
        if st.button(T["logout_btn"], use_container_width=True, type="primary"):
            del st.session_state.final_parsed_rows
            st.session_state.captcha_bytes = None
            if 'input_pw' in st.session_state: del st.session_state['input_pw']
            if 'input_captcha' in st.session_state: del st.session_state['input_captcha']
            cleanup_driver() 
            st.rerun()

components.html(
    """
    <script>
        let attempts = 0;
        let interval = setInterval(function() {
            var parentDoc = window.parent.document;
            var marker = parentDoc.getElementById("main-header-marker");
            if (marker) {
                var headerBlock = marker.closest('div[data-testid="stHorizontalBlock"]');
                if (headerBlock) {
                    headerBlock.classList.add("custom-fixed-header");
                    clearInterval(interval);
                }
            }
            attempts++;
            if(attempts > 10) clearInterval(interval); 
        }, 500);
    </script>
    """,
    height=0,
    width=0
)

st.markdown("<div class='header-spacer'></div>", unsafe_allow_html=True)
st.write(T["desc"])

current_time = time.time()
is_locked_out = current_time < st.session_state.lockout_until
remaining_time = int(st.session_state.lockout_until - current_time)

if not has_data:
    st.subheader(T["login_header"])
    
    col_zh, col_en = st.columns(2)
    with col_zh:
        if st.button("中文", use_container_width=True, disabled=(st.session_state.lang == 'zh')):
            st.session_state.lang = 'zh'
            st.rerun()
    with col_en:
        if st.button("English", use_container_width=True, disabled=(st.session_state.lang == 'en')):
            st.session_state.lang = 'en'
            st.rerun()
            
    st.write("") 
    
    if is_locked_out:
        st.error(T["lockout"].format(time=remaining_time))

    with st.form("login_form"):
        user_id = st.text_input(T["user_id"], placeholder="...", autocomplete="username", key="input_uid")
        user_pw = st.text_input(T["user_pw"], type="password", placeholder="...", autocomplete="current-password", key="input_pw")
        
        if not st.session_state.captcha_bytes:
            submit_action = st.form_submit_button(T["captcha_btn"])
            action_type = "GET_CAPTCHA" if submit_action else None
            captcha_code = ""
        else:
            st.image(st.session_state.captcha_bytes)
            captcha_code = st.text_input(T["captcha_input"], autocomplete="off", key="input_captcha")
            submit_action = st.form_submit_button(T["login_btn"], type="primary", disabled=is_locked_out)
            action_type = "LOGIN" if submit_action else None

    if action_type == "GET_CAPTCHA":
        with st.spinner(T["loading_sys"]):
            try:
                cleanup_driver() 
                driver = init_driver()
                driver.get(LOGIN_URL)
                wait = WebDriverWait(driver, 20)
                captcha_node = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ContentPlaceHolder1_SysCaptchaNdhu img")))
                img_src = captcha_node.get_attribute("src")
                if img_src and "base64," in img_src:
                    st.session_state.captcha_bytes = base64.b64decode(img_src.split("base64,")[1])
                    st.session_state.driver_instance = driver
                    st.rerun()
            except Exception as e:
                cleanup_driver()
                st.error(T["error_timeout"])

    elif action_type == "LOGIN":
        if not user_id or not user_pw or not captcha_code:
            st.warning("Please fill in all fields / 請填寫所有欄位")
        else:
            progress_info = st.empty()
            try:
                driver = st.session_state.driver_instance
                if not driver:
                    raise WebDriverException("Browser instance lost.")
                wait = WebDriverWait(driver, 30)
                progress_info.info(T["auth_ing"])
                
                wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_txtUID"))).send_keys(user_id)
                driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword").send_keys(user_pw)
                driver.find_element(By.ID, "ContentPlaceHolder1_txtCaptcha").send_keys(captcha_code)
                
                current_url = driver.current_url
                login_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_imgbtnLogin")))
                login_btn.click()
                
                try:
                    WebDriverWait(driver, 5).until(EC.url_changes(current_url))
                except TimeoutException:
                    pass 
                    
                if "Login.aspx" in driver.current_url:
                    st.session_state.fail_count += 1
                    if st.session_state.fail_count >= 3:
                        st.session_state.lockout_until = time.time() + 15
                        st.session_state.fail_count = 0
                    else:
                        st.error(T["error_login"].format(count=st.session_state.fail_count))
                    progress_info.empty()
                    cleanup_driver() 
                    st.session_state.captcha_bytes = None
                    st.rerun()
                else:
                    st.session_state.fail_count = 0 
                    progress_info.info(T["fetch_ing"])
                    gpa_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[2]/div[4]/div/ul/li[3]/a')))
                    gpa_btn.click()
                    wait.until(lambda d: len(d.window_handles) > 1)
                    driver.switch_to.window(driver.window_handles[-1])
                    
                    if st.session_state.lang == 'en':
                        try:
                            lang_btn = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_btnSwitchLang")))
                            if lang_btn.get_attribute("value") == "English":
                                lang_btn.click()
                                wait.until(EC.staleness_of(lang_btn)) 
                        except Exception:
                            pass 
                            
                    progress_info.info(T["parsing"])
                    dropdown = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_YearSemeDDList")))
                    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_YearSemeDDList"]/option[@value="0:0"]').click()
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "td[data-th]")))
                    soup = BeautifulSoup(driver.page_source, 'html.parser')
                    parsed_rows = []
                    for tr in soup.find_all("tr"):
                        row_data = {}
                        for td in tr.find_all("td"):
                            head = td.get('data-th')
                            if not head: continue
                            head = head.strip()
                            text = td.get_text(strip=True)
                            if head in ["學年", "Acad. Year"]: row_data['year'] = text
                            elif head in ["學期", "Seme."]: row_data['seme'] = text
                            elif head in ["科目名稱", "Course Title", "Course Name", "Subject"]: row_data['name'] = text
                            elif head in ["必/選修", "Required/Elective"]: row_data['req_elec'] = text
                            elif head in ["學分", "Credit", "Credits"]: 
                                try: row_data['credit'] = float(text)
                                except: row_data['credit'] = 0.0
                            elif head in ["成績", "Grade", "Score"]: row_data['grade'] = text
                        if 'name' in row_data and 'grade' in row_data:
                            parsed_rows.append(row_data)
                            
                    st.session_state.final_parsed_rows = parsed_rows
                    progress_info.empty()
                    
            except Exception as e:
                error_msg = "S_ERR_01" if "password" in str(e).lower() else "S_ERR_02"
                st.error(T["error_process"].format(err=error_msg))
            finally:
                cleanup_driver()
                st.session_state.captcha_bytes = None
                if has_data or 'final_parsed_rows' in st.session_state:
                    st.rerun()

if has_data:
    st.components.v1.html(
        """
        <script>
            function forceScrollToTop() {
                var parent = window.parent;
                if (parent) {
                    parent.scrollTo({ top: 0, behavior: 'smooth' });
                    var mainSection = parent.document.querySelector('section.main') || parent.document.querySelector('.stApp');
                    if (mainSection) {
                        mainSection.scrollTo({ top: 0, behavior: 'smooth' });
                    }
                }
            }
            setTimeout(forceScrollToTop, 100);
            setTimeout(forceScrollToTop, 500);
        </script>
        """,
        height=0,
    )

    parsed_rows = st.session_state.final_parsed_rows
    overall_stats, last4_stats, last60_stats, valid_courses = compute_gpa_analytics(parsed_rows)
    
    def display_metrics_card(header, stats):
        g45, g40, g43, total_c = stats
        with st.container(border=True):
            st.subheader(header)
            cols = st.columns(3)
            cols[0].metric(f"{T['overall_gpa']} (4.5)", f"{g45:.2f}")
            cols[1].metric(f"{T['overall_gpa']} (4.3)", f"{g43:.2f}")
            cols[2].metric(f"{T['overall_gpa']} (4.0)", f"{g40:.2f}")

    with st.expander(T["conv_header"]):
        conv_data = {
            T["conv_grade"]: ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D", "E"],
            T["conv_45_col"]: [4.5, 4.0, 3.7, 3.3, 3.0, 2.7, 2.5, 2.3, 2.0, 1.0, 0.0],
            T["conv_43_col"]: [4.3, 4.0, 3.7, 3.3, 3.0, 2.7, 2.3, 2.0, 1.7, 1.3, 0.0],
            T["conv_40_col"]: [4.0, 4.0, 3.67, 3.33, 3.0, 2.67, 2.33, 2.0, 1.67, 1.33, 0.0]
        }
        df_conv = pd.DataFrame(conv_data)
        df_conv.index += 1 
        st.table(df_conv.style.format({
            T["conv_45_col"]: "{:.2f}",
            T["conv_43_col"]: "{:.2f}",
            T["conv_40_col"]: "{:.2f}"
        }))

    st.write("") 
    display_metrics_card(T["summary_header"], overall_stats)
    display_metrics_card(T["last4_header"], last4_stats)
    display_metrics_card(T["last60_header"], last60_stats)

    st.markdown("---")
    with st.expander(T["detail_tab"]):
        df_display_data = []
        for r in valid_courses[::-1]:
            df_display_data.append([
                f"{r['year']}-{r['seme']}", 
                r['name'], 
                r['req_elec'],
                r['credit'], 
                r['grade']
            ])
            
        df = pd.DataFrame(df_display_data, columns=[T["col_term"], T["col_name"], T["col_req"], T["col_credit"], T["col_grade"]])
        df.index += 1 
        st.dataframe(df, use_container_width=True, height=400)

st.markdown("---")
st.caption(T["privacy"])