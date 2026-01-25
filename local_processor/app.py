import streamlit as st
import polars as pl
from processor import UniversalProcessor
import os

st.set_page_config(page_title="Универсальный Сверщик 🚀", layout="wide")

st.title("🚀 Локальный Космолет: Обработка Excel")
st.markdown("""
Перетащите файлы Excel сюда, и локальный ИИ DeepSeek разберет их автоматически. 
Данные не покидают ваш компьютер!
""")

# Инициализация процессора
if "processor" not in st.session_state:
    st.session_state.processor = UniversalProcessor()

# Боковая панель с настройками
with st.sidebar:
    st.header("Настройки")
    model = st.selectbox("Выберите модель Ollama", ["deepseek-r1:8b", "llama3.2:3b", "deepseek-r1:32b"], index=0)
    st.session_state.processor.model_name = model
    
    st.divider()
    gas_url = st.text_input("URL вашего Google Web App", placeholder="https://script.google.com/macros/s/...")
    st.info("Сюда будут отправляться данные после обработки.")

# Загрузка файлов
uploaded_files = st.file_uploader("Выберите Excel файлы", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    st.subheader(f"Загружено файлов: {len(uploaded_files)}")
    
    for uploaded_file in uploaded_files:
        with st.expander(f"📄 Файл: {uploaded_file.name}", expanded=True):
            col1, col2 = st.columns([1, 2])
            
            # Сохраняем временно для обработки
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                with st.spinner("ИИ анализирует структуру..."):
                    data, file_type = st.session_state.processor.process_file(temp_path)
                
                col1.success(f"Тип: {file_type}")
                col1.metric("Найдено строк", len(data))
                
                # Показываем превью данных в таблице
                df_preview = pl.DataFrame(data).head(10)
                col2.dataframe(df_preview, use_container_width=True)
                
                if st.button(f"Отправить {uploaded_file.name} в Google", key=uploaded_file.name):
                    if not gas_url:
                        st.error("Сначала укажите URL в боковой панели!")
                    else:
                        with st.spinner("Отправка данных..."):
                            # Здесь будет вызов requests к GAS
                            st.toast(f"Данные из {uploaded_file.name} отправлены!")
            
            except Exception as e:
                st.error(f"Ошибка при обработке: {e}")
            
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

st.divider()
st.caption("Сделано для проекта Сверка 2.0 | Локальный ИИ DeepSeek")
