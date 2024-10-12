import time
import logging
import datetime as dt
from Claude_generated import check_expiry,load_holiday_data
from zoneinfo import ZoneInfo

# lets import our telegram methods
from telegram.ext import Application,CommandHandler,ContextTypes
from telegram import Update
from telegram.ext import Updater,CommandHandler
from api_key import api_key


# lets do some basic logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='bot.log',
    level=logging.INFO
)


# defining our global variables
subscribed_users = set()
holiday_dict = load_holiday_data('Book1.xlsx')
expiries = {
        'Midcap Nifty': 'Monday',
        'FINNIFTY': 'Tuesday',
        'BANKNIFTY': 'Wednesday',
        'NIFTY': 'Thursday'
    }

async def start(update:Update,context:ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('Hello! I am your Expiry Tracker Telegram bot.')    


async def track_expiry(update:Update,context:ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    subscribed_users.add(user_id)
    await update.message.reply_text("Thanks for enabling, now you'll get daily updates of what expiry is it today!!")
    logging.info(f"User {user_id} subscribed to expiry updates")

    await send_expiry_update(context)

async def send_expiry_update(context:ContextTypes.DEFAULT_TYPE):
    today = dt.datetime.now().date()
    expiry_info = check_expiry(today,holiday_dict,expiries)


    logging.info(f"Sending expiry update: {expiry_info}")
    
    if not subscribed_users:
        logging.warning("No subscribed users to send update to")
    
    for user_id in subscribed_users:
        try:
            await context.bot.send_message(chat_id=user_id, text=expiry_info)
            logging.info(f"Sent update to user {user_id}")
        except Exception as e:
            logging.error(f"Failed to send message to user {user_id}: {e}")


async def help(update:Update,context:ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('I can remind you of daily Index expiries as well as expiries coming in the following week!!')

async def about(update:Update,context:ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('''***This bot was created by Harsh, to help him track what Index is expiring today in the Indian stock markets!
    He's given me special functionality to remind you at 12AM, 9AM, 2PM and 3PM about today's expiry. 
    While also taking care of any holidays!***''')

def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logging.error(f"Exception while handling an update: {context.error}")


def  main():
    application = Application.builder().token(api_key).build()

    application.add_handler(CommandHandler('start',start))
    application.add_handler(CommandHandler('help',help))
    application.add_handler(CommandHandler('track',track_expiry))
    application.add_handler(CommandHandler('about',about))
    timezone = ZoneInfo('Asia/Kolkata')
    job_queue = application.job_queue
    job_queue.run_daily(send_expiry_update,time=dt.time(hour=0, minute=45,tzinfo=timezone))
    job_queue.run_daily(send_expiry_update,time=dt.time(hour=9, minute=0,tzinfo=timezone))
    job_queue.run_daily(send_expiry_update, time=dt.time(hour=14, minute=0,tzinfo=timezone))
    job_queue.run_daily(send_expiry_update, time=dt.time(hour=15, minute=00,tzinfo=timezone))
    job_queue.run_daily(send_expiry_update, time=dt.time(hour=15, minute=20,tzinfo=timezone))
    # job_queue.run_once(send_expiry_update, 10)

    application.add_error_handler(error_handler)
    application.run_polling()

if __name__=='__main__':
    main()
