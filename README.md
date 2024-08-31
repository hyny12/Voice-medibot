# Voice Medi Bot

Welcome to Voice Medi Bot, a project designed to make it easier for doctors to understand patient feedback. Good communication between patients and doctors is essential for providing the best care. Voice Medi Bot helps by turning spoken feedback into text and analyzing the emotions behind it.

With this tool, doctors can quickly see how patients feel about their care through easy-to-understand star ratings. By using the latest technology, Voice Medi Bot makes the feedback process simpler and more effective, helping to improve patient satisfaction and overall healthcare quality.

## Overview
Voice Medi Bot is a healthcare application that enhances patient-doctor interactions by integrating speech recognition and sentiment analysis. The bot converts spoken words into text using the Google Speech Recognition API and analyzes the sentiment using TextBlob. The analyzed data is then stored in Excel files, and sentiment scores are converted into star ratings for easy feedback interpretation.

## Features
- **Speech Recognition:** Converts spoken words into text using Google Speech Recognition API.
- **Sentiment Analysis:** Analyzes the sentiment of patient feedback using TextBlob.
- **Star Ratings:** Converts sentiment analysis scores into star ratings for simplified feedback.
- **Data Storage:** Saves analyzed data in Excel files for record-keeping and analysis.
- **User-Friendly Interface:** Provides a straightforward and intuitive interface for healthcare providers.

## Installation
To set up the Voice Medi Bot on your local machine, follow these steps:

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/voice-medi-bot.git
    ```

2. Navigate to the project directory:
    ```bash
    cd voice-medi-bot
    ```

3. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

4. Set up the Google Speech Recognition API key:
    - Obtain an API key from the [Google Cloud Console](https://console.cloud.google.com/).
    - Set up the key in your environment or in the script.

5. Run the application:
    ```bash
    python main.py
    ```

## Usage
1. Launch the application by running `main.py`.
2. Speak into the microphone when prompted.
3. The bot will convert your speech into text and analyze the sentiment.
4. The sentiment score and corresponding star rating will be displayed, and the data will be saved in an Excel file.

## Contributing
We welcome contributions to improve the Voice Medi Bot! To contribute:

1. Fork the repository.
2. Create a new branch for your feature or bugfix:
    ```bash
    git checkout -b feature/your-feature-name
    ```
3. Commit your changes:
    ```bash
    git commit -m "Add new feature"
    ```
4. Push to your branch:
    ```bash
    git push origin feature/your-feature-name
    ```
5. Open a Pull Request and describe your changes.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact
For any inquiries, feel free to contact:
- **Hyny** - hyny.s7@gmail.com
- LinkedIn: [Your LinkedIn Profile](https://www.linkedin.com/in/hyny-s-101a04228/)
