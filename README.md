Automation for Secretaria Municipal da Cultura Brasil/São Paulo 🎭

This repository contains an automation project developed for the Municipal Department of Culture. The main goal is to facilitate and streamline the sending of messages via WhatsApp to coordinators and logistics groups, based on information contained in control spreadsheets.

Features 🚀
Data Reading: The code reads information from two Excel spreadsheets: "Controle De Logistica.xlsx" and "Geral Macro .xlsm".
Personalized Message Sending: Messages are customized according to the data in each row of the spreadsheets.
Interaction with WhatsApp Web: The automation simulates interactions with WhatsApp Web to send the messages.
Delay Between Sends: There's a mechanism to delay between sends to avoid WhatsApp blocking.

Requirements 🛠️
Python 3.x
Libraries: pandas, time, pyautogui, and pywhatkit
Spreadsheets "Controle De Logistica.xlsx" and "Geral Macro .xlsm" with correct data.

Usage 📋
Make sure you have Python installed in your environment.
Install the required libraries by executing pip install pandas time pyautogui pywhatkit.
Clone this repository to your local machine.
Ensure you have the spreadsheets "Controle De Logistica.xlsx" and "Geral Macro .xlsm" in the project folder.
Run the script automacao_secretaria.py.

Additional Documentation 📑
For details about variables, execution flow, and other specific information about the code, refer to the internal comments.

Contribution 🤝
Contributions are welcome! Feel free to open an issue or submit a pull request with improvements, fixes, or new features.

Author 🧑‍💻
This project was developed by João Veloso.

