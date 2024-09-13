# Automated-Reporting-System

## **Disclaimer:**
> **I am unable to provide the complete code and detailed data analysis due to the project being under a Non-Disclosure Agreement (NDA) with the respective company. The data used in this analysis contains sensitive company information, and therefore, cannot be published or shared to maintain confidentiality and the integrity of the company’s data.**

## Bot Explained

## 1. WhatsApp Bot

This bot is designed to automate interactions with WhatsApp Web using Selenium and manage data via Google Sheets through an API.

### Key Components:
- **Selenium**: Automates interactions with WhatsApp Web, allowing the bot to open, search, and send messages automatically.
- **Google Sheets (gspread)**: Integrates with Google Sheets to read and write data, which is then used for message scheduling and delivery.
- **Tabulate**: Structures and displays data from Google Sheets in table format for easy viewing.
- **ServiceAccountCredentials**: Authenticates access to Google Sheets via a JSON credentials file.

### Main Functions:
- **show_data()**: Retrieves and displays selected data from Google Sheets in a table format.
- **show_data2()**: Similar to show_data(), but fetches different rows or data sets.
- **Selenium Automation**: The bot logs into WhatsApp Web, scans the QR code, and sends messages based on the data retrieved.
- **Automated Messaging**: Messages can be sent to contacts or groups based on a set schedule or user input.

### Benefits and Applications:
- Ideal for businesses that need automated mass messaging or scheduled notifications.
- Saves time by automating the process of copying data from Google Sheets and sending it through WhatsApp.

---

## 2. Telegram Bot

Built using python-telegram-bot, this bot automates data handling and reporting via Telegram while integrating with Google Sheets.

### Key Components:
- **python-telegram-bot**: Handles interactions between the bot and users, processing commands like `/start`, `/data`, and more.
- **gspread & Google Sheets API**: Similar to WhatsApp bot, used for reading and writing data directly from Google Sheets.
- **schedule**: Automates tasks such as sending reports or fetching data at specific times.

### Main Functions:
- **show_data(), show_data2(), show_data3(), show_data_report()**: Fetches and displays data from Google Sheets based on user commands.
- **record_data()**: Allows users to submit new data via Telegram, which is then recorded in Google Sheets.
- **Scheduled Tasks**: The bot can automatically fetch data or send reports at specified intervals.

### Command Management:
- Uses handlers to process commands like `/data`, `/record`, and more.
- Provides instant feedback to users, such as showing data tables or confirming successful data entries.

### Benefits and Applications:
- Useful for organizations looking to manage data or provide automated updates via messaging apps.
- Perfect for tasks like team data collection, automatic report generation, or real-time notifications.

---

## Conclusion:
Both the WhatsApp and Telegram bots are designed to enhance efficiency in managing data and automating communications. The WhatsApp bot leverages Selenium to automate interactions with WhatsApp Web, while the Telegram bot uses the Telegram API to handle user commands. Integration with Google Sheets allows both bots to automate data processing, making them powerful tools for businesses needing flexible, automated reporting and messaging systems.

---

## Bot Workflow

1. **Start**: The system initiates and enters standby mode.
2. **System Standby**: The system waits for a specified time or condition to trigger further actions.
3. **Time Check (8 AM to 10 PM)**: The system verifies if the current time is between 8 AM and 10 PM.
   - If "No", it returns to standby mode.
   - If "Yes", it proceeds to the next check.
4. **Minute Check (O’clock)**: The system checks if the current time is at the beginning of an hour (minute 00).
   - If "No", it returns to standby.
   - If "Yes", it proceeds to read and copy data.
5. **Read Data and Copy**: The system fetches and copies relevant data.
6. **Send Data to Messaging Platform**: The system sends the copied data to messaging platforms such as WhatsApp or Telegram.
7. **Data Limit Check**: The system checks if the inputted data has reached its daily limit.
   - If "No", it continues inputting data.
   - If "Yes", it creates a new row or updates the date.
8. **Create New Row/Update Date**: If the daily limit is reached, the system creates a new row or updates the date in the source.
