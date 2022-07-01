def get_plus2(number):
    if number < 0:
        return '+'+str(number * (-1))
    return str(number*-1)
print(get_plus2(100))