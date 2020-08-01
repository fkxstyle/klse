def string_value_converter(text):
    if text[-1:] == 'K' or text[-1:] == 'k':  # Check if the last digit is K
        value = float(text[:-1]) * 1000  # Remove the last digit with [:-1], and convert to int and multiply by 1000
    elif text[-1:] == 'M' or text[-1:] == 'm':  # Check if the last digit is M
        value = float(text[:-1]) * 1000000  # Remove the last digit with [:-1], and convert to int and multiply by 1000000
    elif text[-1:] == '%':  # Check if the last digit is %, percentage
        value = float(text[:-1])  # Remove the last digit with [:-1], and convert to float
    else:  # just in case data doesnt have an M or K
        value = float(text)
    return value