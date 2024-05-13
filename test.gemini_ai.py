import google.generativeai as genai
# from google.colab import userdata

GOOGLE_API_KEY='AIzaSyAPRgK8jXW2gUFvzqnyIW7qoLFejB2z4Ic'
genai.configure(api_key=GOOGLE_API_KEY)

for m in genai.list_models():
  if 'generateContent' in m.supported_generation_methods:
    print(m.name)