import streamlit as st
import requests

# Configuration de la page (titre dans l’onglet, largeur, etc.)
# st.set_page_config(
#    page_title="Boîte à outils Executive Relocations",
#    layout="wide",
#    page_icon="🏠",  # Optionnel : emoji ou chemin vers une image
#)

# Titre principal
#st.title("🏠 Accueil")
st.title("Bienvenue sur la boîte à outils d'Executive Relocations")

# Message d’explication
st.info("← Sélectionnez une application dans le menu de gauche pour commencer.")



# ---------------------------------------
# CONFIGURATION
# ---------------------------------------

API_KEY = "81ab95cf7b136e129d510a0e9f09bac2"  # Remplacer par ta clé OpenWeatherMap
VILLE = "Gennevilliers"
URL = f"https://api.openweathermap.org/data/2.5/weather?q={VILLE}&appid={API_KEY}&units=metric&lang=fr"

# ---------------------------------------
# FONCTION METEO
# ---------------------------------------
def get_weather():
    response = requests.get(URL)
    if response.status_code == 200:
        data = response.json()
        temperature = round(data["main"]["temp"])
        description = data["weather"][0]["description"].capitalize()
        icon_code = data["weather"][0]["icon"]
        icon_url = f"http://openweathermap.org/img/wn/{icon_code}@2x.png"
        return temperature, description, icon_url
    else:
        return None, None, None

# ---------------------------------------
# AFFICHAGE STREAMLIT
# ---------------------------------------
st.subheader(f"🌤️ Météo à {VILLE}")

temperature, description, icon_url = get_weather()

if temperature is not None:
    col1, col2 = st.columns([1, 3])

    with col1:
        st.image(icon_url, width=80)

    with col2:
        st.write(f"**{description}**")
        st.write(f"**Température : {temperature}°C**")
else:
    st.error("Impossible de récupérer la météo pour le moment.")


