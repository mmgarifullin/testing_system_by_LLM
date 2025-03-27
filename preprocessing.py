import os
import json
import time
import base64
import subprocess
import webbrowser
import streamlit as st
import concurrent.futures
from phi.document.reader.pdf import PDFReader_1
import io
import numpy as np
import pandas as pd
from streamlit import session_state as ss
from openpyxl import load_workbook
from client import Client, Config


class WebInterface(object):

    def __init__(self):
        self.users = self.load_users()
        self.preprocessing = None
        if "_dir_path" not in ss:
            ss._dir_path = os.path.abspath(
                os.path.dirname(__file__))
        if "client" not in ss:
            ss.client = Client.launch()
        self.setup_ui()
        

    def load_users(self):
        if os.path.exists("users.json"):
            with open("users.json", "r") as f:
                return json.load(f)
        return {}

    def save_users(self):
        with open("users.json", "w") as f:
            json.dump(self.users, f)

    @staticmethod
    def get_base64(image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()

    @staticmethod
    def save_to_json(data, directory, filename):
        os.makedirs(directory, exist_ok=True)

        file_path = os.path.join(directory, filename)
        with open(file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data, json_file, ensure_ascii=False, indent=4)

    @staticmethod
    def open_rag_work():
        try:
            result = subprocess.run(r"C:\Users\user\run_ollama.bat", shell=True, check=True)
            print(f"Команда выполнена успешно: {result}")
        except subprocess.CalledProcessError as e:
            print(f"Ошибка: {e}")
            print(f"Код возврата: {e.returncode}")
            print(f"Вывод: {e.output}")
        try:
            result_1 = subprocess.run([r"C:\Users\user\lss.exe"], check=True)
        except subprocess.CalledProcessError as e:
            print(f"Ошибка при запуске: {e}")
        webbrowser.open("http://127.0.0.1:5555")

    def setup_ui(self):
        st.set_page_config(page_title="Программа интеллектуальной автоматизированной оценки эссе-тестов")

        # image_path = "C:/Users/user/Downloads/n_5da9914b5aff2.jpg"
        image_path = os.path.join(ss._dir_path, "background.png")

        st.markdown(
            f"""
        <style>
        .stApp {{
            background-image: url("data:image/jpeg;base64,{self.get_base64(image_path)}");
            background-size: cover;
            background-position: center;
            height: 100vh;
            color: red !important;  
            font-size: 20px !important;
        }}
        .stButton {{
            font-size: 20px !important; /* Увеличиваем размер шрифта */
            font-family: 'Courier New', Courier, monospace; /* Пример военного шрифта */
            font-weight: bold; /* Жирный текст */
            color: red; /* Цвет текста на кнопках */
            text-align: center; /* Выравнивание текста по центру */
            display: inline-block; /* Отображение в строку */
            margin: 1px 1px; /* Отступы между кнопками */
            cursor: pointer; /* Курсор при наведении */
            width: 30px; 
            height: 80%;
        }}
            /* Hide the Streamlit header and menu */
            header {{visibility: hidden;}}
            /* Optionally, hide the footer */
            .streamlit-footer {{display: none;}}
        </style>
        """,
            unsafe_allow_html=True
        )

        with st.sidebar:
            st.markdown(
                "<h1 style='text-align: center; color: grey;'>Программа автоматизированной интеллектуальной оценки ответов тест-эссе</h1>", unsafe_allow_html=True)
            # st.image(r"C:/Users/user/Desktop/sv-big.jpg")
            st.image(os.path.join(ss._dir_path, "sv-big.jpg"))

            if st.button("Регистрация в системе", use_container_width=True):
                st.session_state.current_action = 'register'
            elif st.button("Вход в систему", use_container_width=True):
                st.session_state.current_action = 'login'
            elif st.button("Загрузить материалы", use_container_width=True):
                st.session_state.current_action = 'upload'
            elif st.button("Тестирующая система", use_container_width=True):
                st.session_state.current_action = "api"

        if 'current_action' in st.session_state:
            action = st.session_state.current_action
            if action == 'register':
                self.register_user()
            elif action == 'login':
                self.login_user()
            elif action == 'upload':
                self.upload_files()
            elif action == 'api':
                self.api_call()

    def register_user(self):

        st.markdown(
            """
        <style>
        .stTextInput>div>input {
            font-size: 40px !important; 
            color: red !important;
        }
        .stTextInput>label {
            font-size: 40px !important;  
            color: red !important;
        }
        </style>
        """,
            unsafe_allow_html=True
        )
        username_admin = None
        password_admin = None

        username_admin = st.text_input("Введите логин администратора:", key="admin_username",
                                       value=st.session_state.get("admin_username", ""))
        password_admin = st.text_input("Введите пароль администратора:", type='password',
                                       key="admin_password", value=st.session_state.get("admin_password", ""))
        if username_admin == 'admin' and password_admin == 'admin':
            username = st.text_input("Введите логин для нового пользователя:", key="reg_username",
                                     value=st.session_state.get("reg_username", ""))
            password = st.text_input("Введите пароль для нового пользователя :", type='password',
                                     key="reg_password", value=st.session_state.get("reg_password", ""))
        username = None
        password = None

        st.markdown(
            """
            <style>
            .fixed-button {
                width: 200px !important;
                height: 50px !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <style>
            .stButton[data-baseweb="button"][data-testid="stButton_confirm_button1"] > button,
            .stButton[data-baseweb="button"][data-testid="stButton_confirm_button2"] > button {
                width: 200px !important;
                height: 50px !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        if st.button("Далее", key='confirm_button1'):
            if username_admin == 'admin' and password_admin == 'admin':
                if st.button("Зарегистрироваться", key='confirm_button2'):
                    if username and password:
                        if username not in self.users:
                            self.users[username] = password
                            self.save_users()  # Сохраняем пользователей после регистрации
                            st.success("Пользователь зарегистрирован!")
                        else:
                            st.error("Пользователь с таким логином уже существует.", icon="🚨")
                    else:
                        st.error("Пожалуйста, введите логин и пароль.", icon="🚨")
            else:
                st.error("Введены неправильные данные.", icon="🚨")

    def login_user(self):

        st.markdown(
            """
        <style>
        .stTextInput>div>input {
            font-size: 40px !important;  
            color: red !important;
        }
        .stTextInput>label {
            font-size: 40px !important;  
            color: red !important;
        }
        </style>
        """,
            unsafe_allow_html=True
        )

        username = st.text_input("Введите логин:", key="login_username", value=st.session_state.get("login_username", ""))
        password = st.text_input("Введите пароль:", type='password', key="login_password",
                                 value=st.session_state.get("login_password", ""))

        st.markdown(
            """
        <style>
        confirm_button3 {
            width: 200px !important;
            height: 50px !important;
        }
        </style>
        """,
            unsafe_allow_html=True,
        )

        if st.button("Войти", key='confirm_button3'):
            if username in self.users and self.users[username] == password:
                self.open_rag_work()
                st.success("Вход выполнен успешно!")
            else:
                st.error("Неверный логин или пароль.", icon="🚨")

    def upload_files(self):
        # Инициализация состояния для логина и пароля
        if 'upload_username' not in st.session_state:
            st.session_state.upload_username = ""
        if 'upload_password' not in st.session_state:
            st.session_state.upload_password = ""
        if 'uploaded_files' not in st.session_state:
            st.session_state.uploaded_files = None
        if 'process_files' not in st.session_state:
            st.session_state.process_files = False
        if 'is_authenticated' not in st.session_state:
            st.session_state.is_authenticated = False
        st.markdown(
            """
        <style>
        .stTextInput>div>input {
            font-size: 40px !important;  
            color: red !important;
        }
        .stTextInput>label {
            font-size: 40px !important;  
            color: red !important;
        }
        .stNumberInput>div>input {
            font-size: 20px !important;  /* Уменьшаем размер шрифта */
            color: red !important;
        }
        .stNumberInput>label {
            font-size: 20px !important;  /* Уменьшаем размер шрифта */
            color: red !important;
        }
        </style>
        """,
            unsafe_allow_html=True
        )
        # Поля для ввода логина и пароля
        username = st.text_input("Введите логин для доступа к обработке файлов:",
                                 key="upload_username", value=st.session_state.upload_username)
        password = st.text_input("Введите пароль для доступа к обработке файлов:", type='password',
                                 key="upload_password", value=st.session_state.upload_password)

        # Добавление PDF в базу знаний
        if "file_uploader_key" not in st.session_state:
            st.session_state["file_uploader_key"] = 100

        uploaded_file = st.sidebar.file_uploader(
            "Прикрепите PDF :page_facing_up:", type="pdf", key=st.session_state["file_uploader_key"], accept_multiple_files=True
        )

        # Виджеты для ввода целых чисел
        col1, col2 = st.columns(2)
        col3, col4 = st.columns(2)

        with col1:
            fragment_add = st.number_input("Введите номер фрагмента для добавления:",
                                           min_value=1, value=1, key="fragment_add")

        with col2:
            fragment_delete = st.number_input("Введите номер фрагмента для удаления:",
                                              min_value=1, value=1, key="fragment_delete")

        st.markdown(
            """
        <style>
        .confirm_button4 {
            width: 200px !important;
            height: 50px !important;
        }
        </style>
        """,
            unsafe_allow_html=True,
        )

        if st.button("Подтвердить", key='confirm_button4'):
            if username in self.users and self.users[username] == password:
                st.session_state.is_authenticated = True
                st.success("Успешная аутентификация!")
        file_names = []
        with col3:
            if st.button("Добавить фрагмент в базу знаний"):
                if st.session_state.is_authenticated:
                    reader = PDFReader_1(chunk_size=200, separators=['/t'])
                    documents = []
                    for doc in uploaded_file:
                        doc_data = reader.read(doc, doc.name[:-4], fragment_add)
                        if doc_data:
                            documents.append(doc_data)
                            st.success(f"Документ '{doc.name}' прикреплён")

                            output_dir_pdf = r'C:/Users/user/data/source'
                            file_name = doc.name[:-4] + '_' + str(fragment_add) + ".pdf"
                            file_names.append(file_name)
                            file_path = os.path.join(output_dir_pdf, file_name)

                            with open(file_path, "wb") as f:
                                f.write(doc.getbuffer())

                    all_fragments = []
                    for doc_fragments in documents:
                        if not doc_fragments:
                            continue
                        all_fragments.extend(doc_fragments)  # Добавляем все фрагменты в общий список

                    # Записываем все фрагменты в один JSON файл
                    if all_fragments:
                        output_dir = r'C:/Users/user/data/text'
                        title = doc_fragments[0]['title'] + '_' + str(fragment_add)
                        filename = f"{title}.json".replace('/', '_')

                        with open(os.path.join(output_dir, filename), 'w', encoding='utf-8') as json_file:
                            json.dump(all_fragments, json_file, ensure_ascii=False, indent=4)

                    st.success(f"Фрагмент {fragment_add} добавлен в базу знаний.")
                    st.session_state.process_files = True

        with col4:
            if st.button("Удалить фрагмент из базы знаний"):
                output_dir_one = r'C:/Users/user/data/text'
                output_dir_two = r'C:/Users/user/data/index'

                with concurrent.futures.ThreadPoolExecutor() as executor:
                    futures = []

                    for file_name in os.listdir(output_dir_one):
                        if file_name.endswith('.json'):
                            file_path = os.path.join(output_dir_one, file_name)

                    # Читаем содержимое файла
                    with open(file_path, 'r', encoding='utf-8') as f:
                        my_list = json.load(f)

                        # Фильтруем список, исключая элементы с нужным значением page
                        my_list = [obj for obj in my_list if obj['page'] != fragment_delete]

                        # Записываем обратно только если список изменился
                    if len(my_list) == 0:
                        os.remove(file_path)  # Удаляем файл, если он пустой
                    else:
                        with open(file_path, 'w', encoding='utf-8') as f:
                            json.dump(my_list, f, ensure_ascii=False, indent=4)

                    # Удаляем соответствующий файл .safetensors
                    safetensors_path = os.path.join(output_dir_two, file_name.replace('.json', '.safetensors'))
                    if os.path.exists(safetensors_path):
                        futures.append(executor.submit(os.remove, safetensors_path))

                # Ожидаем завершения всех потоков
                concurrent.futures.wait(futures)
                st.success("Фрагменты успешно удалены из базы знаний.")

    def api_call(self):
        st.markdown(
            """
        <style>
        div[data-testid="InputInstructions"] > span:nth-child(1) {
            visibility: hidden;
        }
        .stMainBlockContainer{
            min-width: 80%;
            max-width: 400rem;
        }
        .stTextInput{
            margin-top: 0.5rem;
            font-size: large;
        }
        .stTextArea{
            margin-top: -1rem;
        }
        .stTextArea>label{
            display: none;
        }
        .stTextArea>div>div>textarea{
            padding: 0 0.5rem;
        }
        div[data-testid="stTextInputRootElement"]{
            background-color: rgba(0,0,0,0);
            border-color: rgba(1,1,1,1);
            &>div{
                background-color: rgba(192,192,192,1);
            }
        }
        .stTextInput>label{
            display: none;
        }
        .stVerticalBlock{
            gap: 0.25rem;
        }
        .stTextInput>div>input {
            font-size: 40px !important;  
            color: red !important;
        }
        .stTextInput>label {
            font-size: 40px !important;  
            color: red !important;
        }
        .stNumberInput>div>input {
            font-size: 20px !important;  /* Уменьшаем размер шрифта */
            color: red !important;
        }
        .stNumberInput>label {
            font-size: 20px !important;  /* Уменьшаем размер шрифта */
            color: red !important;
        }
        div[data-testid="stCheckbox"]>label>div>div{
            color: red;
        }
        </style>
        """,
            unsafe_allow_html=True
        )
        src_col, _, dst_col = st.columns(
            (8, 1, 8), gap="small", vertical_alignment="bottom")
        if 'src_tex' not in ss:
            ss.src_tex = [""]
        if 'src_lbl' not in ss:
            ss.src_lbl = ["Эталонный ответ на 1 вопрос"]
        if 'dst_tex' not in ss:
            ss.dst_tex = [""]
        if 'dst_lbl' not in ss:
            ss.dst_lbl = ["Ответ слушателя на 1 вопрос"]
        if 'result' not in ss:
            ss.result = ""
        if 'show_all' not in ss:
            ss.show_all = False
        if 'params' not in ss:
            ss.params = Config.load().asdict()
        if 'listener_name' not in ss:
            ss.listener_name = ""

        def add_src():
            ss.src_lbl.append(f"Эталонный ответ на {len(ss.src_tex) + 1} вопрос")
            ss.src_tex.append("")

        def rem_src():
            if len(ss.src_lbl) == 1:
                return
            ss.src_lbl.pop()
            ss.src_tex.pop()

        def add_dst():
            ss.dst_lbl.append(f"Ответ слушателя на {len(ss.dst_tex) + 1} вопрос")
            ss.dst_tex.append("")

        def rem_dst():
            if len(ss.dst_lbl) == 1:
                return
            ss.dst_lbl.pop()
            ss.dst_tex.pop()

        def assign_grade(score, boundaries):
            if score >= boundaries['5']:
                return 5
            elif score >= boundaries['4']:
                return 4
            elif score >= boundaries['3']:
                return 3
            else:
                return 2

        name = st.text_input("Введите фамилию слушателя:", key="listener_surname")

        

        def call_api():

            ok, result = ss.client(ss.src_tex, ss.dst_tex, ss.params)
            
            if ok and ss.show_all:
                mir = pd.MultiIndex.from_tuples([
                    (i, j) for i in result.keys()
                    for j in ss.dst_lbl], names=["Тип", "Ответ"])
                mic = pd.Index([i for i in ss.src_lbl])
                dat = result["score"]
                for k in ("dense", "sparse", "colbert"):
                    dat.extend(result[k])
                df = pd.DataFrame(
                    dat,
                    columns=mic,
                    index=mir)
                ss.result = df
            elif ok:

                df = pd.DataFrame(
                    result["score"],
                    columns=ss.src_lbl,
                    index=ss.dst_lbl)
                diagonal_elements = np.diag(df.values)
                diagonal_elements = np.array([round(elem, 2) for elem in diagonal_elements])
                new_df = pd.DataFrame([diagonal_elements], index = [f"{name}"], columns=df.T.columns)
                average_score = new_df.mean(axis=1)
                new_df['Средний балл'] = average_score
                new_df['Оценка'] = new_df['Средний балл'].apply(lambda score: assign_grade(score, {'5':ss.params['5'], '4':ss.params['4'], '3':ss.params['3'], '2':ss.params['2']}))
                ss.result = new_df
                file_path = r'C:/Users/user/output.xlsx'
                with pd.ExcelWriter(file_path) as writer:
                    new_df.to_excel(writer, sheet_name='Лист1')
            else:
                ss.result = result
        
        
        if st.button("Скачать результат тестирования в формате XLSX", use_container_width=True):
            #save_to_excel(ss.result)
            if hasattr(ss, 'result') and isinstance(ss.result, pd.DataFrame):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    ss.result.to_excel(writer, index=True)
                    output.seek(0)
                st.download_button(
                label="Скачать",
                data=output,
                file_name="result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
                )

        with src_col:
            for k, (l, t) in enumerate(zip(ss.src_lbl, ss.src_tex)):
                ss.src_lbl[k] = st.text_input("_", value=l, key=f"sl_{k}")
                ss.src_tex[k] = st.text_area("_", value=t, key=f"st_{k}")
            src_col.button(r"Добавить эталонный ответ на вопрос", on_click=add_src,
                           use_container_width=True)
            src_col.button(r"Удалить эталонный ответ на вопрос", on_click=rem_src,
                           use_container_width=True)
        with dst_col:
            for k, (l, t) in enumerate(zip(ss.dst_lbl, ss.dst_tex)):
                ss.dst_lbl[k] = st.text_input("_", value=l, key=f"dl_{k}")
                ss.dst_tex[k] = st.text_area("_", value=t, key=f"dt_{k}")
            dst_col.button(r"Добавить ответ слушателя", on_click=add_dst,
                           use_container_width=True)
            dst_col.button(r"Удалить ответ слушателя", on_click=rem_dst,
                           use_container_width=True)

        settings, result = st.columns((8, 2), gap="small")
        with settings:
            with st.popover("Настройки", use_container_width=True):
                if st.checkbox("Показать всё"):
                    ss.show_all = True
                else:
                    ss.show_all = False
                ss.params['use_mult'] = st.checkbox(
                    "Использовать для расчёта умножение вместо сложения",
                    value=ss.params['use_mult'])
                ss.params['dense_coeff'] = st.number_input(
                    "Коэффициент семантической компоненты",
                    min_value=0., value=ss.params['dense_coeff'])
                ss.params['sparse_coeff'] = st.number_input(
                    "Коэффициент компоненты соответствия по словам", min_value=0., value=ss.params['sparse_coeff'])
                ss.params['colbert_coeff'] = st.number_input(
                    "Коэффициент расширенной семантической компоненты",
                    min_value=0., value=ss.params['colbert_coeff'])
                ss.params['atten_coeff'] = st.number_input(
                    "Экспонента аттенюации результатов",
                    min_value=0., value=ss.params['atten_coeff'])
                ss.params['divider'] = st.number_input(
                    "Знаменатель в формуле суммации компонент",
                    min_value=1., value=ss.params['divider'])
                ss.params['5'] = st.number_input('Нижняя граница 5:', value=90.0, min_value=0.0, max_value=100.0, step=1.0)
                ss.params['4'] = st.number_input('Нижняя граница 4:', value=80.0, min_value=0.0, max_value=100.0, step=1.0)
                ss.params['3'] = st.number_input('Нижняя граница 3:', value=50.0, min_value=0.0, max_value=100.0, step=1.0)
                ss.params['2'] = st.number_input('Нижняя граница 2:', value=0.0, min_value=0.0, max_value=100.0, step=1.0)
        with result:
            result.button("Проверка", on_click=call_api, use_container_width=True)
        with st.container():
            ss.result


if __name__ == "__main__":
    web_interface = WebInterface()
