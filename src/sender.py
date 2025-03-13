import pywhatkit as wp

def send_whatsapp_message(phone_no, message, wait_time=15): 
    wp.sendwhatmsg_instantly(
        phone_no, 
        message, 
        wait_time=wait_time, 
        tab_close=True, 
        close_time=3
    )
    return True