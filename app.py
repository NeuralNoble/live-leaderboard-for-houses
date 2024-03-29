from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)


class House:
    def __init__(self, name, logo_filename, points):
        self.name = name
        self.logo_filename = logo_filename
        self.points = points


# Function to read data from an Excel file and convert it into a DataFrame
def read_excel_to_dataframe(file_path):
    try:
        df = pd.read_excel(file_path, index_col='Event')
        return df
    except FileNotFoundError:
        print("File not found.")
        return None
    except Exception as e:
        print("An error occurred:", e)
        return None


# Function to rank the houses based on the points in the 'total' row
def rank_houses(data_frame):
    try:
        total_row = data_frame.loc['total'].dropna()  # Drop NaN values
        ranked_houses = total_row.sort_values(ascending=False)
        return ranked_houses
    except KeyError:
        print("Row named 'total' not found.")
        return None
    except Exception as e:
        print("An error occurred:", e)
        return None


# Route to display the leaderboard
@app.route('/')
def leaderboard():
    # Read data from Excel file
    excel_file_path = "leaderboard.xlsx"  # Replace with your file path
    data_frame = read_excel_to_dataframe(excel_file_path)

    if data_frame is not None:
        # Rank the houses
        ranked_houses = rank_houses(data_frame)
        if ranked_houses is not None:
            # Get top three houses
            top_three = ranked_houses.head(3)
            # Create House objects and populate them with data
            houses = []
            for house_name, points in top_three.items():
                # Generate filename for logo
                logo_filename = f"{house_name.lower()}_logo.png"
                house = House(house_name, logo_filename, points)
                houses.append(house)
            return render_template('leaderboard.html', top_three=houses, house_points=data_frame)

    # If data loading or ranking fails, return an empty response
    return ""


if __name__ == '__main__':
    app.run(debug=True)