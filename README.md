import requests
import random
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches
import os
import google.generativeai as genai


# ====== جزء جلب الصور من Pexels ======
def fetch_new_image_from_pexels(api_key, query):
    """
    Fetch a new, random image from Pexels API.

    Args:
        api_key (str): The Pexels API key for authentication.
        query (str): The search query for the image.

    Returns:
        The image content in bytes and the image URL.
    """
    api_url = "https://api.pexels.com/v1/search"
    headers = {"Authorization": api_key}

    # عشوائية الصفحة من 1 إلى 50
    random_page = random.randint(1, 50)
    params = {"query": query, "per_page": 10, "page": random_page}

    response = requests.get(api_url, headers=headers, params=params)
    if response.status_code == 200:
        data = response.json()
        photos = data.get("photos", [])
        if photos:
            # اختيار صورة عشوائية من الصور الموجودة
            random_photo = random.choice(photos)
            image_url = random_photo["src"]["original"]
            image_response = requests.get(image_url)
            if image_response.status_code == 200:
                return image_response.content, image_url
            else:
                print("Failed to download the image.")
        else:
            print("No images found for the query.")
    else:
        print(f"Failed to fetch data from Pexels API. Status code: {response.status_code}")
    return None, None


# ====== جزء كتابة النص باستخدام Google Generative AI ======
def generate_script(api_key):
    """
    Generates a short poetic script using Google Generative AI.

    Args:
        api_key: The API key for Google Generative AI.

    Returns:
        The generated script text.
    """
    genai.configure(api_key=api_key)
    prompt = "اكتب تعليقاً غزلياً على صورة."
    model = genai.GenerativeModel('gemini-1.5-flash')
    chat = model.start_chat(history=[])
    response = chat.send_message(prompt)
    return response.text


# ====== جزء حفظ الصورة والنص في ملف Word ======
def save_image_and_script_to_word(image_content, script_text, filename="output.docx"):
    """
    Saves the image and its script to a Word document.

    Args:
        image_content: The binary content of the image.
        script_text: The script text to save.
        filename: The name of the Word document.
    """
    # إنشاء مستند جديد أو فتح مستند موجود
    doc = Document()

    # حفظ الصورة مؤقتًا
    temp_image_path = "temp_image.jpg"
    with open(temp_image_path, "wb") as img_file:
        img_file.write(image_content)

    # إضافة الصورة والنص إلى المستند
    doc.add_picture(temp_image_path, width=Inches(5))
    doc.add_paragraph(script_text)

    # حفظ المستند
    doc.save(filename)
    print(f"Image and script saved to {filename}")

    # إزالة ملف الصورة المؤقت
    os.remove(temp_image_path)


# ====== الجزء الرئيسي ======
if __name__ == "__main__":
    # مفاتيح API
    PEXELS_API_KEY = "uDwf6Aywwf1B6Whz1nXrgFLEC2XIlIxsX8oScnmTc4GQ5c6RW25hpIBM"
    GOOGLE_API_KEY = "AIzaSyBtBQh4e-ZZnmEoGGkWO0iU08Ote-7OF6k"

    # البحث عن الصور
    query = "nature"
    image_content, image_url = fetch_new_image_from_pexels(PEXELS_API_KEY, query)

    if image_content:
        # كتابة النص باستخدام الذكاء الاصطناعي
        script = generate_script(GOOGLE_API_KEY)

        # حفظ الصورة والنص في ملف Word
        save_image_and_script_to_word(image_content, script)
    else:
        print("Failed to fetch an image.")


