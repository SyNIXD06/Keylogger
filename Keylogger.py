import msoffcrypto
import string
from itertools import product
import time
import openpyxl

charset = string.printable
maxrange = 18
encrypted = open("contraseñas.xlsx", "rb")
file = msoffcrypto.OfficeFile(encrypted)

def solve_password(maxrange):
    isCracked = False
    passwords = []
    for i in range(0, maxrange + 1):
        print("Loop #: ", i)
        if isCracked:
            break
        for attempt in product(charset, repeat=i):
            tmpPass = ''.join(attempt)
            try:
                file.load_key(tmpPass)  
                print("Password correct: ", attempt)  # Array
                print("password: ", tmpPass)  # String
                passwords.append(tmpPass)
                isCracked = True
                break
            except:
                print("Exception opening the file, password incorrect: ", attempt)
    return passwords

start_time = time.time()
print("charset: ", charset)
cracked_passwords = solve_password(maxrange)
execution_time = time.time() - start_time
print("--- %s seconds ---" % execution_time)
print("--- %s min ---" % (execution_time / 60))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Cracked Passwords"

for index, password in enumerate(cracked_passwords):
    ws.cell(row=index + 1, column=1, value=password)

wb.save("cracked_passwords.xlsx")
print("Resultados guardados en cracked_passwords.xlsx")

