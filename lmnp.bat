@echo off
title Lancement LMNP
echo Activation de l'environnement et lancement de Streamlit...

:: 1. Se déplacer dans le dossier du projet (Remplacez le chemin ci-dessous)
cd /d "C:\Users\Amine\Desktop\Formulaire LMNP"

:: 2. Lancer streamlit
py -m streamlit run lmnp.py

pause