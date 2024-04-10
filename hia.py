import streamlit as st
from streamlit import session_state as ss
import pandas as pd
import io
from loguru import logger
import streamlit_authenticator as stauth

from report_utils import create_docx_report, generate_report_data, show_named_plotly_colours
import yaml
from yaml.loader import SafeLoader
st.set_page_config(layout="wide")

def next_page():
    placeholder.empty()
    if(ss['page'] < PAGES):
        ss['page'] += 1

def previous_page():
    placeholder.empty()
    if(ss['page'] > 1):
        ss['page'] -= 1


with open('config.yaml') as file:
    config = yaml.load(file, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
    config['pre-authorized']
)
cols1, cols2, cols3 = st.columns([1,1,1])
with cols2:
    fields = {'Form name':'Đăng nhập', 'Username':'Tên đăng nhập', 'Password':'Mật khẩu',
                      'Login':'Login'}
    authenticator.login(fields=fields)


if st.session_state["authentication_status"]:
    authenticator.logout()
    st.write(f'Welcome *{st.session_state["name"]}*')
    # st.title('Some content')

    PAGES = 6

    report_data = {}
    cate_df = pd.read_csv('category.csv')
    report_data['cate_df'] = cate_df
    
    st.sidebar.image('logo.png')
    # st.sidebar.markdown(f"<h1 style='text-align: center;'>Chương Trình</h1>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<h2 style='text-align: center;'>Phân Tích Sự Cố Y Khoa</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("""---""")



    # st.sidebar.write("")
    # st.sidebar.write("")


    if 'page' not in ss:
        ss['page'] = -1

    uploaded_file = st.sidebar.file_uploader("Chọn file dữ liệu:", type = ['xlsx','csv'])

    st.sidebar.markdown("""---""")

    #TODO get avaliable year from data file
    avaliable_year = [2023,2024]
    report_year = st.sidebar.selectbox('Chọn năm báo cáo', options=avaliable_year, index=0)

    avaliable_term = [
        'Báo cáo tổng kết năm',
        'Báo cáo 6 tháng đầu năm',
        'Báo cáo 6 tháng cuối năm'
    ]
    term = st.sidebar.radio('Chọn kỳ báo cáo', options=avaliable_term, index=0)
    selected_term_idx = avaliable_term.index(term)

    from_month = 7 if selected_term_idx==2 else 1
    to_month = 6 if selected_term_idx==1 else 12

    term_year = f'năm {report_year}' if selected_term_idx==0\
                    else f'6 tháng đầu năm {report_year}' if selected_term_idx==1\
                    else f'6 tháng cuối năm {report_year}'

    report_data['term_year'] = term_year


    if not uploaded_file:
        ss['page'] = -1
    else:
        try:
            df = pd.read_excel(uploaded_file, index_col=0)
            df['Thời gian'] = pd.to_datetime(df['Thời gian'], format = '%d/%m/%Y')
            report_data.update(generate_report_data(df, report_year, from_month, to_month))
            if 'page' not in ss or ss['page'] in [-1,0]:
                ss['page'] = 1
        except Exception as e:
            logger.exception(e)
            ss['page'] = 0


    placeholder = st.empty()

    with st.empty():
        if(ss['page'] == -1):
            # show_named_plotly_colours()
            st.markdown(f"<h2 style='text-align: center;'>Chưa chọn file dữ liệu báo báo</h2>", unsafe_allow_html=True)

        if(ss['page'] == 0):
            st.markdown(f"<h2 style='text-align: center;'>File bị lỗi hoặc không có dữ liệu trong kỳ!</h2>", unsafe_allow_html=True)

        if(ss['page'] == 1):
            with st.container():
                st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
                st.subheader(f'Danh sách sự cố y khoa {term_year}')
                styler_1 = report_data['styler_1']
                st.write(styler_1)
                st.markdown("""---""")
                cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])

                with cols5:
                    st.button('Trang kế tiếp',key='page1next',on_click=next_page)

        if(ss['page'] == 2):
            with st.container():
                st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
                st.subheader(f'Biểu đồ phát hiện sự cố theo thời gian xuất hiện')

                styler_2 = report_data['styler_2']
                st.markdown(styler_2.hide().to_html().replace('<td>', '<td align="right">'), unsafe_allow_html=True)
                fig_2 = report_data['fig_2']
                st.plotly_chart(fig_2)

                st.markdown("""---""")
                cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])
                with cols1:
                    st.button('Trang trước',key='page2pre',on_click=previous_page)

                with cols5:
                    st.button('Trang kế tiếp',key='page2next',on_click=next_page)

        if(ss['page'] == 3):
            with st.container():
                st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
                st.subheader(f'Biểu đồ phát hiện sự cố theo địa điểm xuất hiện')

                styler_3 = report_data['styler_3']
                st.markdown(styler_3.hide().to_html().replace('<td>', '<td align="right">'), unsafe_allow_html=True)
                fig_3 = report_data['fig_3']
                st.plotly_chart(fig_3)

                st.markdown("""---""")
                cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])
                with cols1:
                    st.button('Trang trước',on_click=previous_page)
                with cols5:
                    st.button('Trang kế tiếp',on_click=next_page)

        if(ss['page'] == 4):
            with st.container():
                st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
                st.subheader(f'Biểu đồ phát hiện sự cố theo mức độ nguy hại cho người bệnh')

                styler_4 = report_data['styler_4']
                st.markdown(styler_4.hide().to_html().replace('<td>', '<td align="right">'), unsafe_allow_html=True)
                fig_4 = report_data['fig_4']
                st.plotly_chart(fig_4)

                st.markdown("""---""")
                cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])
                with cols1:
                    st.button('Trang trước',on_click=previous_page)
                with cols5:
                    st.button('Trang kế tiếp',on_click=next_page)

        if(ss['page'] == 5):
            with st.container():
                st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
                st.subheader(f'Biểu đồ phát hiện sự cố theo nhóm sự cố')

                styler_5 = report_data['styler_5']
                st.markdown(styler_5.hide().to_html().replace('<td>', '<td align="right">'), unsafe_allow_html=True)
                fig_5 = report_data['fig_5']
                st.plotly_chart(fig_5)

                st.markdown("""---""")
                cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])
                with cols1:
                    st.button('Trang trước',on_click=previous_page)
                with cols5:
                    st.button('Trang kế tiếp',on_click=next_page)

        # if(ss['page'] == 6):
        #     with st.container():
        #         st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
        #         st.subheader(f'Nhận xét và kiến nghị')

        #         comment = ss['comment'] if 'comment' in ss else ''
        #         comment = st.text_area(
        #             "Nhận xét:",
        #             comment,
        #             height=None, max_chars=None, key=None)
        #         if comment!="":
        #             ss['comment'] = comment
        #             report_data['comment']=comment

        #         st.markdown("""---""")
        #         recommendation = ss['recommendation'] if 'recommendation' in ss else ''
        #         recommendation = st.text_area(
        #             "Kiến nghị:",
        #             recommendation,
        #             height=None, max_chars=None, key=None)
        #         if recommendation!="":
        #             ss['recommendation'] = recommendation
        #             report_data['recommendation']=recommendation

        #         st.markdown("""---""")
        #         cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])
        #         with cols1:
        #             st.button('Trang trước',on_click=previous_page)
        #         with cols5:
        #             st.button('Trang kế tiếp',on_click=next_page)

        if(ss['page'] == 6):
            with st.container():
                st.markdown(f"<h2 style='text-align: center;'>Phân tích sự cố y khoa dạng biểu đồ {term_year}</h2>", unsafe_allow_html=True)
                st.subheader(f'Tải về kết quả phân tích')
                
                doc_download = create_docx_report(report_data)

                bio = io.BytesIO()
                doc_download.save(bio)
                if doc_download:
                    st.download_button(
                        label="Download",
                        data=bio.getvalue(),
                        file_name="report.docx",
                        mime="docx"
                    )

                st.markdown("""---""")
                cols1, cols2, cols3, cols4, cols5 = st.columns([1,1,1,1,1])
                with cols1:
                    st.button('Trang trước',on_click=previous_page)

elif st.session_state["authentication_status"] is False:
    cols1, cols2, cols3 = st.columns([1,1,1])
    with cols2:
        st.error('Tên đăng nhập hoặc mật khẩu không đúng')
elif st.session_state["authentication_status"] is None:
    cols1, cols2, cols3 = st.columns([1,1,1])
    with cols2:
        st.warning('Xin vui lòng nhập tên đăng nhập và mật khẩu')