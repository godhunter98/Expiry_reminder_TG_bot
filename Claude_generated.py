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

def get_next_working_day(date, holiday_dict):
    # Start with the day after the given date
    next_day = date + timedelta(days=1)
    # Keep incrementing the day until we find a non-holiday
    while is_holiday(next_day, holiday_dict):
        next_day += timedelta(days=1)
    return next_day

def check_expiry(date, holiday_dict, expiries):
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
        tomorrow_expiring_indices = [index for index, expiry_day in expiries.items() if expiry_day == tomorrow_day_name]
        
        # If there are indices expiring tomorrow, list them as expiring today due to the holiday
        if tomorrow_expiring_indices:
            message.append(f"Indices expiring today ({date.strftime('%B %d, %Y')}) due to tomorrow's holiday:")
            for index in tomorrow_expiring_indices:
                message.append(f"- {index}")
    
    # Check for today's regular expiries
    if not is_holiday(original_date,holiday_dict):
        day_name = date.strftime('%A')
        # Find indices that are scheduled to expire today
        today_expiring_indices = [index for index, expiry_day in expiries.items() if expiry_day == day_name]
        # Print today's regular expiries, if any
        if today_expiring_indices:
            message.append(f"\nRegularly scheduled expiring indices for today ({date.strftime('%B %d, %Y')}, {day_name}):")
            for index in today_expiring_indices:
                message.append(f"- {index}")
        else:
            message.append(f"\nNo regularly scheduled indices are expiring today ({date.strftime('%B %d, %Y')}, {day_name})")
    else: 
        holiday_info = holiday_dict[pd.Timestamp(original_date)]
        message.append(f"\nNo regularly scheduled indices are expiring today ({date.strftime('%B %d, %Y')}, cause there is a market holiday on account of: {holiday_info['Share Market Holiday 2024']} )")

    return '\n'.join(message)

def main():
    # Load holiday data from the Excel file
    holiday_dict = load_holiday_data('Book1.xlsx')
    
    # Define the expiry schedule for different indices
    expiries = {
        'Midcap Nifty': 'Monday',
        'FINNIFTY': 'Tuesday',
        'BANKNIFTY': 'Wednesday',
        'NIFTY': 'Thursday'
    }

    # Get today's date
    today = datetime.today().date()-timedelta(5)
    
    # Check expiries for today, considering tomorrow's potential holiday
    check_expiry(today, holiday_dict, expiries)


# Ensure that the script is being run directly (not imported as a module)
if __name__ == "__main__":
    main()