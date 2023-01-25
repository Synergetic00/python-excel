'''Excel function in Python'''

def ABS(number: str | int) -> int:
    '''Returns the absolute value of a number'''
    try: number = int(str(number))
    except: raise TypeError('Only integers are allowed')
    return abs(number)
