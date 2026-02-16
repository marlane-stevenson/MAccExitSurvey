import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import re

# Set style for plots
sns.set_theme(style="whitegrid")

def parse_columns(columns):
    """
    Parses column names to identify course type, category, and course name.
    Returns a dictionary mapping column name to metadata.
    """
    col_map = {}

    # Categories we care about for ranking
    rank_categories = ["Most Beneficial", "Neutral", "Least Beneficial"]
    # "Did not take" is also a category but we filter it out usually

    for col in columns:
        if "MAcc CORE courses" in col:
            c_type = "CORE"
        elif "MAcc Elective courses" in col:
            c_type = "ELECTIVE"
        else:
            continue

        if " - Ranks - " not in col:
            continue

        if not col.endswith(" - Rank"):
            continue

        # Extract category
        # Pattern: ... - Ranks - {Category} - {Course} - Rank
        # We can try to find the category in the string
        found_cat = None
        for cat in rank_categories + ["Did not take"]:
            if f" - Ranks - {cat} - " in col:
                found_cat = cat
                break

        if not found_cat:
            print(f"Warning: Could not identify category in column: {col}")
            continue

        # Extract course name
        # It is between " - Ranks - {cat} - " and " - Rank"
        prefix = f" - Ranks - {found_cat} - "
        suffix = " - Rank"

        start_idx = col.find(prefix) + len(prefix)
        end_idx = col.rfind(suffix)

        if start_idx == -1 or end_idx == -1:
            continue

        course_name = col[start_idx:end_idx]

        col_map[col] = {
            "type": c_type,
            "category": found_cat,
            "course": course_name
        }

    return col_map

def process_rankings(df, col_map, course_type):
    """
    Calculates the average global rank for courses of a given type.
    """
    # Filter columns for this type
    type_cols = [c for c, meta in col_map.items() if meta["type"] == course_type]

    # We need to process per respondent (row)
    course_ranks = [] # List of (Course, Rank) tuples from all students

    # Dictionary to hold list of ranks for each course
    course_rank_lists = {}

    for idx, row in df.iterrows():
        # Collect all course responses for this student
        student_courses = []

        for col in type_cols:
            meta = col_map[col]
            val = row[col]

            if pd.isna(val):
                continue

            cat = meta["category"]
            course = meta["course"]

            if cat == "Did not take":
                continue

            # Category Priority (Lower is better)
            # Most Beneficial: 1
            # Neutral: 2
            # Least Beneficial: 3

            cat_priority = 0
            if cat == "Most Beneficial":
                cat_priority = 1
            elif cat == "Neutral":
                cat_priority = 2
            elif cat == "Least Beneficial":
                cat_priority = 3

            # Add to list
            student_courses.append({
                "course": course,
                "cat_priority": cat_priority,
                "rank_val": val
            })

        # Sort student courses to determine global rank
        # Sort by Category Priority, then by Rank Value
        student_courses.sort(key=lambda x: (x["cat_priority"], x["rank_val"]))

        # Assign global rank (1-based)
        for i, item in enumerate(student_courses):
            global_rank = i + 1
            course_name = item["course"]

            if course_name not in course_rank_lists:
                course_rank_lists[course_name] = []

            course_rank_lists[course_name].append(global_rank)

    # Calculate statistics
    results = []
    for course, ranks in course_rank_lists.items():
        avg_rank = sum(ranks) / len(ranks)
        results.append({
            "Course": course,
            "Average Rank": avg_rank,
            "Count": len(ranks)
        })

    results_df = pd.DataFrame(results)
    if not results_df.empty:
        results_df = results_df.sort_values("Average Rank")

    return results_df

def create_plot(df, title, filename):
    if df.empty:
        print(f"No data for {title}")
        return

    plt.figure(figsize=(10, 8))
    # Reset index to get a clean sequential index for the y-axis if we want
    # But sns.barplot can handle it.

    # We want the best rank (lowest number) at the top.
    # So we sort descending for the plot so that the barplot (which plots bottom-up) shows
    # the "smallest average rank" at the top?
    # Or we can just invert the y-axis?

    # Let's sort so the best (lowest rank) is first in the dataframe.
    # We already sorted by Average Rank ascending.

    # Create the bar plot
    ax = sns.barplot(x="Average Rank", y="Course", data=df, color="skyblue")

    plt.title(title, fontsize=16)
    plt.xlabel("Average Rank (Lower is Better)", fontsize=12)
    plt.ylabel("")
    plt.tight_layout()

    # Save
    output_path = os.path.join("outputs", filename)
    plt.savefig(output_path)
    print(f"Saved plot to {output_path}")
    plt.close()

def main():
    data_path = "data/data.xlsx"
    if not os.path.exists(data_path):
        print(f"Error: {data_path} not found.")
        return

    print(f"Reading {data_path}...")
    try:
        df = pd.read_excel(data_path)
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    print("Parsing columns...")
    col_map = parse_columns(df.columns)

    # Process CORE
    print("Processing CORE courses...")
    core_df = process_rankings(df, col_map, "CORE")
    print(core_df)
    create_plot(core_df, "MAcc CORE Course Rankings", "core_rank_order.png")

    # Process ELECTIVE
    print("Processing ELECTIVE courses...")
    elective_df = process_rankings(df, col_map, "ELECTIVE")
    print(elective_df)
    create_plot(elective_df, "MAcc ELECTIVE Course Rankings", "elective_rank_order.png")

    # Save CSVs (optional but helpful)
    core_df.to_csv("outputs/core_rankings.csv", index=False)
    elective_df.to_csv("outputs/elective_rankings.csv", index=False)

if __name__ == "__main__":
    main()
