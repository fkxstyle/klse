import numpy as np

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

def best_fit_slope(xs,ys):
    from statistics import mean
    m = (((mean(xs)*mean(ys)) - mean(xs*ys)) /
         ((mean(xs)**2) - mean(xs**2)))
    return m

def prepare_numeric_value_format(value):
    numeric_type_tuple = (np.int64, np.float64, int, float)
    if isinstance(value, numeric_type_tuple):
        value = '{:,}'.format(round(value,2))
    return value