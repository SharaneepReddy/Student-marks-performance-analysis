import pyodbc
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

server_name = 'DESKTOP-SA5VKMA\\SQLEXPRESS'
database_name = 'Student_Mark_Analysis'

conn = pyodbc.connect('DRIVER={ODBC driver 17 for SQL Server} ;\
                        SERVER=' + server_name + ' ;\
                        DATABASE=' + database_name + ' ;\
                        Trusted_Connection=yes; ')

cursor = conn.cursor()


cursor.execute("""
    SELECT Name, AVG(marks.marks_obtained) AS average_marks
    FROM Student
    JOIN marks ON Student.Student_id = marks.student_id
    GROUP BY Name
    ORDER BY average_marks DESC
""")

results = cursor.fetchall()

print("\nStudent Average Marks:")
averages = {}
for row in results:
    print(f"{row.Name}: {row.average_marks:.2f}")
    averages[row.Name] = row.average_marks

plt.figure(figsize=(12, 6))
names = list(averages.keys())
avg_marks = list(averages.values())
plt.bar(names, avg_marks, color='skyblue')
plt.xticks(rotation=45, ha='right')
plt.title('Average Marks by Student')
plt.ylabel('Average Marks')
plt.tight_layout()
plt.show()

cursor.execute("""
    WITH StudentAverages AS (
        SELECT student.name, AVG(marks.marks_obtained) AS average_marks
        FROM Student
        JOIN marks ON Student.Student_id = marks.student_id
        GROUP BY Name
    )
    SELECT 'Top Performer' AS performance_type, name, average_marks 
    FROM StudentAverages
    WHERE average_marks = (SELECT MAX(average_marks) FROM StudentAverages)
""")

top_performers = cursor.fetchall()

cursor.execute("""
    WITH StudentAverages AS (
        SELECT student.name, AVG(marks.marks_obtained) AS average_marks
        FROM Student
        JOIN marks ON Student.Student_id = marks.student_id
        GROUP BY Name
    )
    SELECT 'Low Performer' AS performance_type, name, average_marks 
    FROM StudentAverages
    WHERE average_marks = (SELECT MIN(average_marks) FROM StudentAverages)
""")

low_performers = cursor.fetchall()

print("\nTop and Low Performers:")
performers = top_performers + low_performers
for row in performers:
    print(f"{row.performance_type}: {row.name} ({row.average_marks:.2f})")

plt.figure(figsize=(8, 5))
types = [row.performance_type for row in performers]
names = [row.name for row in performers]
marks = [row.average_marks for row in performers]
colors = ['green' if t == 'Top Performer' else 'red' for t in types]
plt.bar(names, marks, color=colors)
plt.title('Top and Low Performers')
plt.ylabel('Average Marks')
plt.tight_layout()
plt.show()

query = """
    SELECT 
        s.Name,
        MAX(CASE WHEN sub.subject_name = 'Mathematics' THEN m.marks_obtained END) AS Math,
        MAX(CASE WHEN sub.subject_name = 'Physics' THEN m.marks_obtained END) AS Physics,
        MAX(CASE WHEN sub.subject_name = 'Chemistry' THEN m.marks_obtained END) AS Chemistry
    FROM 
        Student s
    JOIN 
        marks m ON s.Student_id = m.student_id
    JOIN 
        subjects sub ON m.subject_id = sub.subject_id
    WHERE
        sub.subject_name IN ('Mathematics', 'Physics', 'Chemistry')
    GROUP BY 
        s.Name
    ORDER BY 
        s.Name
"""

cursor.execute(query)
results = cursor.fetchall()

print("\nSTUDENT MARKS")
print("-" * 50)
print("{:<20} {:<10} {:<10} {:<10}".format("Student Name", "Math", "Physics", "Chemistry"))
print("-" * 50)
    
subject_data = {'Math': [], 'Physics': [], 'Chemistry': []}
student_names = []
for row in results:
    print("{:<20} {:<10} {:<10} {:<10}".format(
        row.Name, 
        row.Math or "-", 
        row.Physics or "-", 
        row.Chemistry or "-"
    ))
    student_names.append(row.Name)
    subject_data['Math'].append(row.Math or 0)
    subject_data['Physics'].append(row.Physics or 0)
    subject_data['Chemistry'].append(row.Chemistry or 0)

print("-" * 50)

plt.figure(figsize=(14, 7))
bar_width = 0.25
index = np.arange(len(student_names))

plt.bar(index, subject_data['Math'], bar_width, label='Mathematics')
plt.bar(index + bar_width, subject_data['Physics'], bar_width, label='Physics')
plt.bar(index + bar_width*2, subject_data['Chemistry'], bar_width, label='Chemistry')

plt.xlabel('Students')
plt.ylabel('Marks')
plt.title('Subject-wise Marks Comparison')
plt.xticks(index + bar_width, student_names, rotation=45, ha='right')
plt.legend()
plt.tight_layout()
plt.show()
# Grade distribution section (fixed)
grade_query = """
WITH GradeData AS (
    SELECT 
        CASE 
            WHEN (m.marks_obtained * 100.0 / m.total_marks) >= 90 THEN 'A+'
            WHEN (m.marks_obtained * 100.0 / m.total_marks) >= 80 THEN 'A'
            WHEN (m.marks_obtained * 100.0 / m.total_marks) >= 70 THEN 'B'
            WHEN (m.marks_obtained * 100.0 / m.total_marks) >= 60 THEN 'C'
            WHEN (m.marks_obtained * 100.0 / m.total_marks) >= 50 THEN 'D'
            ELSE 'F'
        END AS grade
    FROM 
        marks m
)
SELECT 
    grade, 
    COUNT(*) AS count
FROM 
    GradeData
GROUP BY 
    grade
ORDER BY 
    CASE grade
        WHEN 'A+' THEN 1
        WHEN 'A' THEN 2
        WHEN 'B' THEN 3
        WHEN 'C' THEN 4
        WHEN 'D' THEN 5
        WHEN 'F' THEN 6
    END
"""

# Execute the query and create a DataFrame
cursor.execute(grade_query)
grade_results = cursor.fetchall()
grade_df = pd.DataFrame(grade_results, columns=['grade', 'count'])

# Grade distribution visualization
plt.figure(figsize=(10, 6))
grade_order = ['A+', 'A', 'B', 'C', 'D', 'F']
colors = ['#27ae60', '#2ecc71', '#f39c12', '#e67e22', '#d35400', '#c0392b']  # Gradient from green to red

ax = sns.barplot(x='grade', y='count', data=grade_df, palette=colors, order=grade_order)
plt.title('Grade Distribution', fontsize=16, pad=20)
plt.xlabel('Grade', fontsize=12)
plt.ylabel('Number of Students', fontsize=12)

# Add count labels on top of bars
for p in ax.patches:
    ax.annotate(f'{int(p.get_height())}', 
                (p.get_x() + p.get_width() / 2., p.get_height()),
                ha='center', va='center', 
                xytext=(0, 5), 
                textcoords='offset points')

plt.grid(axis='y', linestyle='--', alpha=0.6)
plt.tight_layout()
plt.show()

cursor.close()
conn.close()