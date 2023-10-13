
def lun_val():
    user_value_digits = []
    user_value = input("Insert number to validate ")
    if user_value.isdigit():

        for index, digit in enumerate(user_value):
            digit = int(digit)
            if index % 2 == 1:
                digit *= 2
                if digit > 9:
                    big_d = decomp_number(digit)
                    user_value_digits.extend(big_d)  # Use extend to add multiple digits
                else:
                    user_value_digits.append(digit)  # Append the doubled digit
            else:
                user_value_digits.append(digit)  # Append the original digit
        print(user_value_digits)

def decomp_number(digit):
    big_number = [int(d) for d in str(digit)]
    summ = sum(big_number)
    return summ

lun_val()