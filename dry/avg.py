# Import necessary libraries
import pandas as pd
from flask import Flask, render_template

# Initialize Flask app
app = Flask(__name__)

# Function to calculate column averages
# Function to calculate average of the last column
def calculate_last_column_average(excel_file):
    # Read Excel file
    df = pd.read_excel(excel_file)
    # Get the last column and calculate its average
    last_column_name = df.columns[-1]
    last_column_average = df[last_column_name].mean()
    return last_column_average

# Route to display average of last column
@app.route('/')
def display_last_column_average():
    # Path to your Excel file
    excel_file_path = 'dry/sentiment_analysis.xlsx'
    # Calculate average of last column
    last_column_average = calculate_last_column_average(excel_file_path)
    # Render HTML template with average of last column
    return render_template('averages.html', last_column_average=last_column_average)


if __name__ == '__main__':
    # Run Flask app
    app.run(debug=True)
