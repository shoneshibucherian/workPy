import os
from SummaryData import insert_batch_headers
from E_Ereport import process_excel_data
from summery import create_summary_report


file_path = 'input.xlsx' 
maintenance_schedule = process_excel_data(file_path)

total_summary = {}
daily_status_dir= os.path.join(os.getcwd(), 'daily_status')
if not os.path.exists(daily_status_dir):
    print("not exists")


file_list=[]
for file in os.listdir(daily_status_dir):
    if file.endswith('.xlsx'):
        file_list.append(os.path.join(daily_status_dir, file))


print("Files in AQAQ daily_status directory:")  
for file in file_list:
    print(file)
    sheet_name=file[file.index("daily_status")+len("daily_status")+1:file.index("Daily")]
    summary_per_day, date=insert_batch_headers(maintenance_schedule,file, "output_batched.xlsx", sheet_name,0)
    total_summary[date] = summary_per_day

print("Total Summary:")
for date, summary in total_summary.items(): 
    # print(f"Date: {date.date()}")
    for key, value in summary.items():
        print(f"{key}: {value}")


create_summary_report(total_summary)