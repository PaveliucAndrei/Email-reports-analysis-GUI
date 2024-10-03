    # Config the main UI element
    MESSAGE_OPENING = ct.CTkLabel(FRAME_EXTRACT, text='', font=('Halvetica', 27))
    MESSAGE_OPENING.place(relx=0.5, rely=0.1, anchor=ct.N)
    BUTTON_MAIN = ct.CTkButton(FRAME_EXTRACT, font=('Helvetica', 22))
    outlook_open = ct.BooleanVar(value=False)

    # Check if the outlook is open
    while outlook_open.get() == False:
        try:
            # Connect to email and extrat the emails list 
            OUTLOOK, Test_XaasIT_emails = email_connection(USER_EMAIL, TEST_XAASIT)
            # OUTLOOK, XaasIT_emails = email_connection(XAASIT_EMAIL, CHECK_ERROR)
            print('All good')
        except AttributeError as err:
            print(err, 'The outlook was not open.')
            # Oppen outlook event
            MESSAGE_OPENING.configure(text='Is Outlook open?')
            BUTTON_MAIN.configure(text='Open Outlook', command=lambda:[open_outlook(), outlook_open.set(True)])
            BUTTON_MAIN.place(relx=0.5, rely=0.2, anchor=ct.N)
            print(outlook_open.get())
            # BUTTON_MAIN.wait_variable(outlook_open)
            
        else:
            # First messages
            MESSAGE_OPENING.configure(text='Select a client')

            # Create the radio buttons
            RADIO_AI = ct.CTkRadioButton(FRAME_EXTRACT, text='AI', value='AI', variable=radio_selection, font=('Helvetica', 22))
            RADIO_AI.place(relx=0.37, rely=0.3, anchor=ct.N)
            RADIO_SKB = ct.CTkRadioButton(FRAME_EXTRACT, text='SKB', value='SKB', variable=radio_selection, font=('Helvetica', 22))
            RADIO_SKB.place(relx=0.52, rely=0.3, anchor=ct.N)
            RADIO_EB = ct.CTkRadioButton(FRAME_EXTRACT, text='EB', value='EB', variable=radio_selection, font=('Helvetica', 22))
            RADIO_EB.place(relx=0.72, rely=0.3, anchor=ct.N)
            # Button extract
            BUTTON_MAIN.configure(text='Extract 📎', command=extracting)
            BUTTON_MAIN.place(relx=0.5, rely=0.5, anchor=ct.N)

            MESSAGE_SELECT = ct.CTkLabel(FRAME_EXTRACT, text='', font=('Halvetica', 24))
            MESSAGE_SELECT.place(relx=0.5, rely=0.7, anchor=ct.N)
            MESSAGE_FILE_LOC = ct.CTkLabel(FRAME_EXTRACT, text='', font=('Halvetica', 16))
            MESSAGE_FILE_LOC.place(relx=0.5, rely=0.8, anchor=ct.N)

        ROOT.mainloop()
