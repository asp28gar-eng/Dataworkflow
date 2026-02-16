import pandas as pd
import matplotlib.pyplot as plt
import os

def analyze_rankings():
    # Define file paths
    input_file = 'Grad Program Exit Survey Data 2024 (1).xlsx'
    output_dir = 'outputs'

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Read the Excel file
    # We read without header first to access the question text in the second row
    try:
        df = pd.read_excel(input_file, header=None)
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")
        return

    # Columns L through S correspond to indices 11 through 18 (0-indexed)
    # Row index 1 contains the question text with course names
    # Row index 3 starts the actual data

    col_start = 11
    col_end = 19 # Exclusive

    # Extract course names
    course_names = []
    for col_idx in range(col_start, col_end):
        question_text = str(df.iloc[1, col_idx])
        # Format is usually "... - Course Name"
        if ' - ' in question_text:
            course_name = question_text.split(' - ')[-1].strip()
        else:
            course_name = question_text
        course_names.append(course_name)

    # Extract ranking data
    # Rows from index 3 to the end
    rankings_data = df.iloc[3:, col_start:col_end]

    # Convert to numeric, coercing errors to NaN
    rankings_data = rankings_data.apply(pd.to_numeric, errors='coerce')

    # Calculate average rank for each course
    avg_ranks = rankings_data.mean()

    # Create a DataFrame for results
    results = pd.DataFrame({
        'Course': course_names,
        'Average Rank': avg_ranks.values
    })

    # Sort by Average Rank (Ascending, because 1 is Best)
    results = results.sort_values(by='Average Rank', ascending=True)

    # Save text report
    report_path = os.path.join(output_dir, 'course_rankings.txt')
    with open(report_path, 'w') as f:
        f.write("Average Course Rankings (Lower score is better):\n")
        f.write("=" * 50 + "\n")
        for index, row in results.iterrows():
            f.write(f"{row['Course']}: {row['Average Rank']:.2f}\n")

    print(f"Text report saved to {report_path}")

    # Generate Bar Chart
    plt.figure(figsize=(12, 8))

    # We invert the y-axis in the plot logic so the best (lowest score) is at the top
    # But usually barh plots from bottom to top.
    # To have the best rank at the top, we should plot the first item (lowest score) at the top.
    # If we just plot barh, the first item in dataframe (index 0) is at the bottom.
    # So let's invert the dataframe order for plotting or invert the axis.
    # Let's just invert the axis after plotting.

    bars = plt.barh(results['Course'], results['Average Rank'], color='skyblue')

    plt.xlabel('Average Rank (1 = Best, 8 = Worst)')
    plt.ylabel('Course')
    plt.title('Average Course Rankings by Students')
    plt.grid(axis='x', linestyle='--', alpha=0.7)

    # Invert y-axis to have the best rank (top of the sorted list) at the top of the chart
    plt.gca().invert_yaxis()

    # Add value labels to the bars
    for bar in bars:
        width = bar.get_width()
        plt.text(width + 0.1, bar.get_y() + bar.get_height()/2,
                 f'{width:.2f}', ha='left', va='center')

    plt.tight_layout()

    chart_path = os.path.join(output_dir, 'course_rankings_chart.png')
    plt.savefig(chart_path)
    print(f"Chart saved to {chart_path}")

if __name__ == "__main__":
    analyze_rankings()
