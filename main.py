import os
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time
import webbrowser
import json
import logging
from selenium.webdriver.remote.remote_connection import LOGGER

# Configure logging
LOGGER.setLevel(logging.WARNING)
logging.basicConfig(level=logging.WARNING)

def get_id_number():
    config_file = 'timesheet_config.json'
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                config = json.load(f)
                if 'id_number' in config:
                    return config['id_number']
        except Exception as e:
            print(f"Error reading config: {e}")
    
    while True:
        id_number = input("Please enter your ID number: ").strip()
        if id_number:
            try:
                with open(config_file, 'w') as f:
                    json.dump({'id_number': id_number}, f)
                return id_number
            except Exception as e:
                print(f"Warning: Could not save ID number ({e}), continuing without saving")
                return id_number

def configure_selenium_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--log-level=3')
    service = Service(log_path=os.devnull)
    service.creation_flags = 0x08000000
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def login_and_get_timesheet(id_number, save_html=True, filename='timesheet.html'):
    driver = configure_selenium_driver()
    
    try:
        driver.get('https://www.sdmataclick.com/m/default.aspx')
        
        # Wait for the ID input field to be present and enter the ID
        id_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'txtIdNumber'))
        )
        id_input.send_keys(id_number)
        
        # Click the login button
        login_button = driver.find_element(By.ID, 'btnSubmit')
        login_button.click()
        
        # Wait for the Me span to be present AND contain non-whitespace text
        def name_is_present(driver):
            try:
                element = driver.find_element(By.ID, 'Me')
                # Try different methods to get the text
                text = element.get_attribute('textContent') or element.text
                return bool(text and text.strip())
            except:
                return False
        
        WebDriverWait(driver, 10).until(name_is_present)

        name_element = driver.find_element(By.ID, 'Me')
        name_html = name_element.get_attribute('innerHTML')
        name_text = name_html.split('<br>')[0].strip() if '<br>' in name_html else name_element.text.split('\n')[0].strip()
        name_text = ' '.join(name_text.split())
        
        # Navigate to the timesheet page
        timesheet_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'btnTimesheet'))
        )
        timesheet_link.click()
        
        # Wait for the timesheet to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'mygrid'))
        )
        
        # Find and extract the table
        table_element = driver.find_element(By.ID, 'mygrid')
        table_html = table_element.get_attribute('outerHTML')
        
        if save_html:
            with open(filename, 'w', encoding='utf-8') as file:
                file.write(table_html)
        
        return table_html, name_text
        
    except Exception as e:
        print(f"Error during web automation: {str(e)}")
        return None, None
    finally:
        driver.quit()
        
def parse_timesheet(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    rows = soup.find_all('tr')[1:]  # Skip header row

    data = []
    for row in rows:
        cols = row.find_all('td')
        date = cols[0].text.strip()
        first_in = cols[1].text.strip()
        last_out = cols[2].text.strip()
        ts_in = cols[3].text.strip()
        ts_out = cols[4].text.strip()
        total_hours = cols[5].text.strip()
        
        # Clean data
        first_in = first_in if first_in and first_in != '     ' else None
        last_out = last_out if last_out and last_out != '     ' else None
        ts_in = ts_in if ts_in != '00:00' else None
        ts_out = ts_out if ts_out != '00:00' else None
        total_hours = float(total_hours) if total_hours else 0.0
        
        data.append({
            'Date': date,
            'FirstIn': first_in,
            'LastOut': last_out,
            'ClockIn': ts_in,
            'ClockOut': ts_out,
            'Hours': total_hours
        })

    df = pd.DataFrame(data)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.sort_values('Date', ascending=False)
    return df

def format_hours_minutes(hours, sign=None):
    if pd.isna(hours):
        return ""
    total_minutes = int(round(hours * 60))
    h, m = divmod(total_minutes, 60)
    
    prefix = ""
    if sign is not None:
        prefix = "+" if sign >= 0 else "-"
    
    if h == 0:
        return f"{prefix}{m}min"
    elif m == 0:
        return f"{prefix}{h}h"
    else:
        return f"{prefix}{h}h {m:02d}min"

def calculate_work_duration(row):
    if pd.isna(row['FirstIn']) or pd.isna(row['LastOut']):
        return 0
    try:
        start = datetime.strptime(row['FirstIn'], '%H:%M')
        end = datetime.strptime(row['LastOut'], '%H:%M')
        return (end - start).total_seconds() / 3600
    except:
        return 0

def analyze_timesheet(df, daily_target=9):
    results = {}
    results['daily_target'] = daily_target
    
    # Filter out weekends (Saturday=5, Sunday=6) and days with 0 hours
    df = df[(df['Date'].dt.dayofweek < 5) & (df['Hours'] > 0)]
    
    daily = df.groupby('Date').agg({
        'FirstIn': 'first',
        'LastOut': 'last',
        'Hours': 'sum'
    }).reset_index()

    daily['DayOfWeek'] = daily['Date'].dt.day_name()
    daily['OnTrack'] = daily['Hours'].apply(lambda x: "✅" if x >= daily_target else "❌")
    daily['Difference'] = daily['Hours'] - daily_target
    daily['WorkDuration'] = daily.apply(calculate_work_duration, axis=1)
    
    daily['HoursFormatted'] = daily['Hours'].apply(format_hours_minutes)
    daily['DifferenceFormatted'] = daily['Difference'].apply(lambda x: format_hours_minutes(abs(x), sign=x))
    
    results['daily'] = daily

    if not daily.empty:
        daily['Week'] = daily['Date'].dt.isocalendar().week
        weekly = daily.groupby('Week').agg({
            'Hours': 'sum',
            'Date': 'nunique',
            'OnTrack': lambda x: (x == "✅").sum()
        }).rename(columns={'Date': 'WorkDays', 'OnTrack': 'OnTargetDays'})
        
        weekly['TargetHours'] = weekly['WorkDays'] * daily_target
        weekly['WeeklyDifference'] = weekly['Hours'] - weekly['TargetHours']
        weekly['AvgDailyHours'] = weekly['Hours'] / weekly['WorkDays']
        weekly['OnTargetPercentage'] = (weekly['OnTargetDays'] / weekly['WorkDays']) * 100
        
        weekly['HoursFormatted'] = weekly['Hours'].apply(format_hours_minutes)
        weekly['TargetHoursFormatted'] = weekly['TargetHours'].apply(format_hours_minutes)
        weekly['WeeklyDifferenceFormatted'] = weekly['WeeklyDifference'].apply(lambda x: format_hours_minutes(abs(x), sign=x))
        weekly['AvgDailyHoursFormatted'] = weekly['AvgDailyHours'].apply(format_hours_minutes)
        
        results['weekly'] = weekly
    
    return results

def generate_html_report(results, user_name=None):
    title = "WORK TIMESHEET ANALYSIS"
    if user_name:
        title = f"{user_name.upper()}'S TIMESHEET ANALYSIS"
    
    daily_target = results['daily_target']
    current_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    daily_chart_data = {
        'labels': results['daily']['Date'].dt.strftime('%Y-%m-%d').tolist(),
        'hours': results['daily']['Hours'].tolist(),
        'differences': results['daily']['Difference'].tolist(),
        'target': [daily_target] * len(results['daily'])
    }
    
    weekly_chart_data = {
        'labels': [f"Week {week}" for week in results['weekly'].index.tolist()],
        'hours': results['weekly']['Hours'].tolist(),
        'targets': results['weekly']['TargetHours'].tolist(),
        'diffs': results['weekly']['WeeklyDifference'].tolist()
    }

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Work Timesheet Analysis</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body {{
                font-family: Arial, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
            }}
            h1, h2 {{
                color: #2c3e50;
            }}
            h1 {{
                border-bottom: 2px solid #3498db;
                padding-bottom: 10px;
            }}
            h2 {{
                border-bottom: 1px solid #eee;
                padding-bottom: 5px;
                margin-top: 30px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
            }}
            th, td {{
                padding: 10px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }}
            th {{
                background-color: #f2f2f2;
                font-weight: bold;
            }}
            tr:hover {{
                background-color: #f5f5f5;
            }}
            .positive {{
                color: green;
                font-weight: bold;
            }}
            .negative {{
                color: red;
                font-weight: bold;
            }}
            .chart-container {{
                display: flex;
                flex-wrap: wrap;
                gap: 20px;
                margin: 30px 0;
            }}
            .chart {{
                flex: 1;
                min-width: 400px;
                height: 400px;
                background: white;
                border-radius: 8px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                padding: 15px;
            }}
            .footer {{
                margin-top: 30px;
                font-size: 0.8em;
                color: #777;
                text-align: center;
            }}
            .note {{
                font-style: italic;
                color: #666;
                margin: 10px 0;
            }}
            .hours-cell {{
                font-family: monospace;
            }}
        </style>
    </head>
    <body>
        <h1>{title}</h1>
        <p><strong>Generated on:</strong> {current_date}</p>
        <p><strong>Daily Target:</strong> {format_hours_minutes(daily_target)} (Weekdays Only)</p>
        <p class="note">Note: Days with 0 hours worked are excluded from analysis</p>
    """
    if not results['daily'].empty:
        html += """
        <h2>DAILY SUMMARY</h2>
        <div class="chart-container">
            <div class="chart">
                <canvas id="dailyChart"></canvas>
            </div>
            <div class="chart">
                <canvas id="dailyDiffChart"></canvas>
            </div>
        </div>
        """
        
        daily_df = results['daily'][['Date', 'DayOfWeek', 'FirstIn', 'LastOut', 'HoursFormatted', 'OnTrack', 'DifferenceFormatted']]
        daily_df = daily_df.rename(columns={
            'DayOfWeek': 'Day of Week',
            'FirstIn': 'Arrival Time',
            'LastOut': 'Leaving Time',
            'HoursFormatted': 'Hours',
            'OnTrack': 'Hours Met?',
            'DifferenceFormatted': 'Difference'
        })
        daily_df['Date'] = daily_df['Date'].dt.strftime('%Y-%m-%d')
        daily_df['Difference'] = daily_df['Difference'].apply(lambda x: f'<span class="{"positive" if x.startswith("+") else "negative"}">{x}</span>')
        # Sort by date descending (most recent first)
        daily_df = daily_df.sort_values('Date', ascending=False)
        html += daily_df.to_html(index=False, escape=False, classes='dataframe')
        
        html += f"""
        <script>
            // Daily Hours Chart
            const dailyCtx = document.getElementById('dailyChart').getContext('2d');
            new Chart(dailyCtx, {{
                type: 'bar',
                data: {{
                    labels: {daily_chart_data['labels']},
                    datasets: [
                        {{
                            label: 'Hours Worked',
                            data: {daily_chart_data['hours']},
                            backgroundColor: 'rgba(54, 162, 235, 0.7)',
                            borderColor: 'rgba(54, 162, 235, 1)',
                            borderWidth: 1
                        }},
                        {{
                            label: 'Daily Target',
                            data: {daily_chart_data['target']},
                            type: 'line',
                            borderColor: 'rgba(255, 99, 132, 1)',
                            borderWidth: 2,
                            fill: false,
                            pointRadius: 0
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Hours'
                            }}
                        }}
                    }},
                    plugins: {{
                        title: {{
                            display: true,
                            text: 'Daily Hours vs Target'
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + ' hours';
                                }}
                            }}
                        }}
                    }}
                }}
            }});

            // Daily Difference Chart
            const dailyDiffCtx = document.getElementById('dailyDiffChart').getContext('2d');
            new Chart(dailyDiffCtx, {{
                type: 'bar',
                data: {{
                    labels: {daily_chart_data['labels']},
                    datasets: [{{
                        label: 'Difference from Target',
                        data: {daily_chart_data['differences']},
                        backgroundColor: function(context) {{
                            const value = context.raw;
                            return value >= 0 ? 'rgba(75, 192, 192, 0.7)' : 'rgba(255, 99, 132, 0.7)';
                        }},
                        borderColor: function(context) {{
                            const value = context.raw;
                            return value >= 0 ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)';
                        }},
                        borderWidth: 1
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        y: {{
                            beginAtZero: false,
                            title: {{
                                display: true,
                                text: 'Hours Difference'
                            }}
                        }}
                    }},
                    plugins: {{
                        title: {{
                            display: true,
                            text: 'Daily Difference from Target'
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + ' hours';
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        </script>
        """
    else:
        html += "<p>No workday data found in the timesheet</p>"
    
    if 'weekly' in results and not results['weekly'].empty:
        html += """
        <h2>WEEKLY SUMMARY</h2>
        <div class="chart-container">
            <div class="chart">
                <canvas id="weeklyChart"></canvas>
            </div>
            <div class="chart">
                <canvas id="weeklyDiffChart"></canvas>
            </div>
        </div>
        """
        
        weekly_df = results['weekly'][['WorkDays', 'OnTargetDays', 'OnTargetPercentage', 'TargetHoursFormatted', 'HoursFormatted', 
                                    'AvgDailyHoursFormatted', 'WeeklyDifferenceFormatted']]
        weekly_df = weekly_df.rename(columns={
            'WorkDays': 'Days',
            'OnTargetDays': 'Met Target',
            'HoursFormatted': 'Hours',
            'TargetHoursFormatted': 'Target',
            'WeeklyDifferenceFormatted': 'Difference',
            'AvgDailyHoursFormatted': 'Avg/Day',
            'OnTargetPercentage': '% On Target'
        })
        weekly_df['Difference'] = weekly_df['Difference'].apply(lambda x: f'<span class="{"positive" if x.startswith("+") else "negative"}">{x}</span>')
        weekly_df['% On Target'] = weekly_df['% On Target'].apply(lambda x: f"{x:.1f}%")
        # Sort by week descending (most recent first)
        weekly_df = weekly_df.sort_index(ascending=False)
        html += weekly_df.to_html(escape=False)
        
        html += f"""
        <script>
            // Weekly Hours Chart
            const weeklyCtx = document.getElementById('weeklyChart').getContext('2d');
            new Chart(weeklyCtx, {{
                type: 'bar',
                data: {{
                    labels: {weekly_chart_data['labels']},
                    datasets: [
                        {{
                            label: 'Hours Worked',
                            data: {weekly_chart_data['hours']},
                            backgroundColor: 'rgba(54, 162, 235, 0.7)',
                            borderColor: 'rgba(54, 162, 235, 1)',
                            borderWidth: 1
                        }},
                        {{
                            label: 'Weekly Target',
                            data: {weekly_chart_data['targets']},
                            backgroundColor: 'rgba(255, 99, 132, 0.7)',
                            borderColor: 'rgba(255, 99, 132, 1)',
                            borderWidth: 1
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Hours'
                            }}
                        }}
                    }},
                    plugins: {{
                        title: {{
                            display: true,
                            text: 'Weekly Hours vs Target'
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + ' hours';
                                }}
                            }}
                        }}
                    }}
                }}
            }});

            // Weekly Difference Chart
            const weeklyDiffCtx = document.getElementById('weeklyDiffChart').getContext('2d');
            new Chart(weeklyDiffCtx, {{
                type: 'bar',
                data: {{
                    labels: {weekly_chart_data['labels']},
                    datasets: [{{
                        label: 'Difference from Target',
                        data: {weekly_chart_data['diffs']},
                        backgroundColor: function(context) {{
                            const value = context.raw;
                            return value >= 0 ? 'rgba(75, 192, 192, 0.7)' : 'rgba(255, 99, 132, 0.7)';
                        }},
                        borderColor: function(context) {{
                            const value = context.raw;
                            return value >= 0 ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)';
                        }},
                        borderWidth: 1
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        y: {{
                            beginAtZero: false,
                            title: {{
                                display: true,
                                text: 'Hours Difference'
                            }}
                        }}
                    }},
                    plugins: {{
                        title: {{
                            display: true,
                            text: 'Weekly Difference from Target'
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + ' hours';
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        </script>
        """
        
    html += f"""
    <div class="footer">
        Report generated on {current_date} | Written and designed by Lee Kaplan (and ChatGPT) | V1.2
    </div>
    </body>
    </html>
    """
    
    return html

def main():
    daily_target_hours = 9
    report_filename = 'timesheet_report.html'
    report_path = os.path.abspath(report_filename)          
    id_number = get_id_number()
    
    try:
        print("Logging in and retrieving timesheet...")
        html_content, user_name = login_and_get_timesheet(id_number)
        
        if html_content is None:
            print("Failed to retrieve timesheet data")
            return
        
        df = parse_timesheet(html_content)
        analysis_results = analyze_timesheet(df, daily_target_hours)
        html_report = generate_html_report(analysis_results, user_name)

        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_report)
        print(f"HTML report saved to: {report_path}")
        print("Opening report in browser...")
        webbrowser.open_new_tab(f'file://{report_path}')
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")  
        
if __name__ == "__main__":
    main()