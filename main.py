# Course: PROG8420-21S-Sec1-Programming for Big Data
# Assignment#: Assignment_3
# Creation Date: Jul 8, 2021
# Author: Harshita Singh
# StudentID: 87846434



def passwordInfoExcel():
    # Create load options
    loadOptions = LoadOptions(chyper - code.XLSX)

    # Set original password
    loadOptions.setPassword("1234")
    # Load the Excel file
    wb = Workbook("workbook-encrypted.xlsx", loadOptions)
    # Set password to none
    wb.getSettings().setPassword(None)
    # Save Excel file
    wb.save("workbook-decrypted.xlsx")

    # Load XLSX workbook
    wb = Workbook("workbook.xlsx")
    # Password protect Excel file
    wb.getSettings().setPassword("1234")
    # Encrypt by specifying the encryption type
    wb.setEncryptionOptions(EncryptionType.XOR, 40)
    # Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider)
    wb.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128)
    # Save Excel file
    wb.save("workbook-encrypted.xlsx")


main_window = tkinter.Tk()
main_window.title('Login Application')
main_window.geometry('400x300')
padd = 20
main_window['padx'] = padd
user_input = tkinter.StringVar()
pass_input = tkinter.StringVar()
info_label = tkinter.Label(main_window, text='Login Application')
info_label.grid(row=0, column=0, pady=20)

info_user = tkinter.Label(main_window, text='Username')
info_user.grid(row=1, column=0)
userInput = tkinter.Entry(main_window, textvariable=user_input)
userInput.grid(row=1, column=1)

info_pass = tkinter.Label(main_window, text='Password')
info_pass.grid(row=2, column=0)
passInput = tkinter.Entry(main_window, textvariable=pass_input)
passInput.grid(row=2, column=1)

login_btn = tkinter.Button(text='Login', command=login)
login_btn.grid(row=3, column=1)

main_window.mainloop()