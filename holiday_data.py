import pandas as pd
from datetime import datetime, timedelta

def load_holiday_data(file_path):
    # Read the Excel file containing holiday data
    data = pd.read_excel(file_path)
    # Convert the 'Date' column to datetime objects for easier manipulation
    data['Date'] = pd.to_datetime(data['Date'])
    # Set 'Date' as the index and select only relevant columns, then convert to a dictionary
    return data.set_index('Date')[['Share Market Holiday 2024', 'Day']].to_dict(orient='index')

def is_holiday(date, holiday_dict):
    # Check if the given date is in the holiday dictionary
    pd_date = pd.Timestamp(date)
    return pd_date in holiday_dict

def is_last_thursday(date):
    # Check if the given date is the last Thursday of the month
    next_week = date + timedelta(days=7)
    return date.weekday() == 3 and next_week.month != date.month

def is_last_tuesday(date):
    # Check if the given date is the last Tuesday of the month
    next_week = date + timedelta(days=7)
    return date.weekday() == 1 and next_week.month != date.month

def check_expiry(date, holiday_dict, weekly_expiries, monthly_expiries):
    message = []
    # Store the original date for reference
    original_date = date
    # Get the next day
    next_day = date + timedelta(days=1)
        
    # Check if tomorrow is a holiday
    if is_holiday(next_day, holiday_dict):
        holiday_info = holiday_dict[pd.Timestamp(next_day)]
        message.append(f"\n****Note: {next_day.strftime('%B %d, %Y')} is a market holiday: {holiday_info['Share Market Holiday 2024']}****\n")
        message.append(f"Checking for expiries that would normally occur tomorrow, today ({date.strftime('%B %d, %Y')})\n")
        
        # Get the day name for tomorrow
        tomorrow_day_name = next_day.strftime('%A')
        # Find indices that would normally expire tomorrow
        tomorrow_expiring_indices = [index for index, expiry_day in weekly_expiries.items() if expiry_day == tomorrow_day_name]
        
        # If there are indices expiring tomorrow, list them as expiring today due to the holiday
        if tomorrow_expiring_indices:
            message.append(f"Indices expiring today ({date.strftime('%B %d, %Y')}) due to tomorrow's holiday:")
            for index in tomorrow_expiring_indices:
                message.append(f"- {index}")
    
    # Check for today's regular expiries
    if not is_holiday(original_date, holiday_dict):
        day_name = date.strftime('%A')
        # Find indices that are scheduled to expire today
        today_expiring_indices = [index for index, expiry_day in weekly_expiries.items() if expiry_day == day_name]
        # Print today's regular expiries, if any
        if today_expiring_indices:
            message.append(f"\nRegularly scheduled weekly expiring indices for today ({date.strftime('%B %d, %Y')}, {day_name}):")
            for index in today_expiring_indices:
                message.append(f"- {index}")
        else:
            message.append(f"\nNo regularly scheduled weekly indices are expiring today ({date.strftime('%B %d, %Y')}, {day_name})")
        
        # Check for monthly expiries
        if is_last_thursday(original_date):
            monthly_expiring_indices = [index for index, expiry_day in monthly_expiries.items() if expiry_day == 'Last Thursday']
            if monthly_expiring_indices:
                message.append(f"\nMonthly expiring indices for today ({date.strftime('%B %d, %Y')}, Last Thursday):")
                for index in monthly_expiring_indices:
                    message.append(f"- {index}")
        elif is_last_tuesday(original_date):
            monthly_expiring_indices = [index for index, expiry_day in monthly_expiries.items() if expiry_day == 'Last Tuesday']
            if monthly_expiring_indices:
                message.append(f"\nMonthly expiring indices for today ({date.strftime('%B %d, %Y')}, Last Tuesday):")
                for index in monthly_expiring_indices:
                    message.append(f"- {index}")
    else: 
        holiday_info = holiday_dict[pd.Timestamp(original_date)]
        message.append(f"\nNo regularly scheduled indices are expiring today ({date.strftime('%B %d, %Y')}, cause there is a market holiday on account of: {holiday_info['Share Market Holiday 2024']} )")

    return '\n'.join(message)

def main():
    # Load holiday data from the Excel file
    holiday_dict = load_holiday_data('Book1.xlsx')
    
    # Define the weekly expiry schedule for different indices
    weekly_expiries = {
        'NIFTY50': 'Thursday',
        'Sensex 50': 'Tuesday'
    }
    
    # Define the monthly expiry schedule for different indices
    monthly_expiries = {
        'NIFTY50': 'Last Thursday',
        'Bank Nifty': 'Last Thursday',
        'FinNifty': 'Last Thursday',
        'Midcap Nifty': 'Last Thursday',
        'Nifty Next50': 'Last Thursday',
        'Sensex': 'Last Tuesday',
        'Bankex': 'Last Tuesday',
        'Sensex 50': 'Last Tuesday'
    }

    # Example usage
    today = datetime.now().date()
    print(check_expiry(today, holiday_dict, weekly_expiries, monthly_expiries))

if __name__ == '__main__':
    main()