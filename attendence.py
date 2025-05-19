import pandas as pd

file_path ="Class_Attendance_CSE 315-CSE(201)(221_D1)-Artificial Intelligence.xlsx"
df = pd.read_excel(file_path)
print("           _   _                 _                  _____      _            _       _             ")
print("     /\\  | | | |               | |                / ____|    | |          | |     | |            ")
print("    /  \\ | |_| |_ ___ _ __   __| | ___  ___ ___  | |     __ _| | ___ _   _| | __ _| |_ ___  _ __ ")
print("   / /\\ \\| __| __/ _ \\ '_ \\ / _` |/ _ \\/ __/ _ \\ | |    / _` | |/ __| | | | |/ _` | __/ _ \\| '__|")
print("  / ____ \\ |_| ||  __/ | | | (_| |  __/ (_|  __/ | |___| (_| | | (__| |_| | | (_| | || (_) | |   ")
print(" /_/    \\_\\__|\\__\\___|_| |_|\\__,_|\\___|\\___\\___|  \\_____\\__,_|_|\\___|\\__,_|_|\\__,_|\\__\\___/|_|   ")
total_students = df.shape[0]
print(f"Total students in sheet: {total_students}")

num_students = int(input("Please enter the Number of Students : "))
print("\n.................................................")
df = df.head(num_students)
data_source = input("Enter 'excel'(offline) to load from Excel file or 'gsheet' to load from Google Sheets: ").strip().lower()

if data_source == 'gsheet':
    gsheet_url = input("Paste the Google Sheets CSV link (ending with output=csv): ").strip()
    try:
        df = pd.read_csv(gsheet_url)
    except Exception as e:
        print("Failed to load Google Sheet. Error:", e)
        exit(1)

attendance_data = df.iloc[:, 3:]
def calculate_percentage(row):
   total_classes = (row == 'P').sum() + (row == 'A').sum()
   if total_classes == 0:
       return 0
   present_count = (row == 'P').sum()
   return round((present_count / total_classes) * 100, 2)

df['Attendance'] = attendance_data.apply(calculate_percentage, axis=1)

def assign_marks(pct):
   if pct >= 70:
       return 5
   elif pct >= 60:
       return 4
   elif pct >= 45:
       return 3
   elif pct >=30:
       return 2
   elif pct <= 30:
       return 1
   else:
       return 0

df['Marks'] = df['Attendance'].apply(assign_marks)

print("\nCalculated Attendance Percentage:")
print("No.  Name                         ID                      Percentage                  Marks")
for idx, row in df.iterrows():
   print(f"{idx+1:<4} {row['Student\'s Name']:<40} {row['Student\'s ID']:<20}   {row['Attendance']:<5}%           {row['Marks']:<5}")

count_70 = df[df['Attendance'] >= 70].shape[0]
count_60 = df[(df['Attendance'] >= 60) & (df['Attendance'] < 70)].shape[0]
count_45 = df[(df['Attendance'] >= 45) & (df['Attendance'] < 60)].shape[0]
count_40 = df[(df['Attendance'] >= 30) & (df['Attendance'] < 45)].shape[0]
count_30 = df[df['Attendance'] <= 30].shape[0]

print("\n.....................................................")
print("\nAttendance Percentage (Student Count):")
print("No.  Percentage   Count")
print(f"1.   >= 70%       {count_70}")
print(f"2.   >= 60%       {count_60}")
print(f"3.   >= 45%       {count_45}")
print(f"3.   >= 30%       {count_40}")
print(f"4.   <= 30%       {count_30}")
