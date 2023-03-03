import random
import string
import re
import PySimpleGUI as sg
import pandas as pd
import email.utils

class PasswordGenerator:
    def __init__(self):
        # Layout
        sg.theme('DarkGrey13')
        layout = [[sg.Text('')],
                  [sg.Text('Password Generator', font=('Roboto', 18, 'bold'), justification='center')],
                  [sg.Text('')],
                  [sg.Text('')],
                  [sg.Text('Address', size=(27, 1), font=('Roboto', 12, 'bold'), justification='left')],
                  [sg.Input(key='address', size=(30, 1), font=('Roboto', 12), justification='center')],
                  [sg.Text('')],
                  [sg.Text('Username or E-mail Address', size=(27, 1), font=('Roboto', 12, 'bold'),
                           justification='left')],
                  [sg.Input(key='userEmail', size=(30, 1), font=('Roboto', 12), justification='center')],
                  [sg.Text('')],
                  [sg.Text('Total Characters', size=(22, 1), font=('Roboto', 12, 'bold'), justification='left'),
                   sg.Combo(values=list(range(25)), key='total_chars', default_value=8, font=('Roboto', 12))],
                  [sg.Input(key='-PASSWORD-', size=(30, 1), font=('Roboto', 12))],
                  [sg.Text('')],
                  [sg.Button('Passwordize', size=(15, 1), font=('Roboto', 10, 'bold'),
                             button_color=('white', '#3F3F3F')),
                   sg.Button('Save', size=(15, 1), font=('Roboto', 10, 'bold'), button_color=('white', '#3F3F3F'))],
                  [sg.Text('')],
                  ]

        self.window = sg.Window('Password Generator', layout, element_justification='center', resizable=False)

    def start(self):
        password_generated = False
        while True:
            event, values = self.window.read()
            if event == sg.WINDOW_CLOSED:
                break
            elif event == 'Passwordize':
                self.new_password = self.generate_password(values)
                self.window['-PASSWORD-'].update(self.new_password)
                password_generated = True
                self.window['Save'].update(disabled=False)  # Enable the Save button
            elif event == 'Save':
                if password_generated:
                    saved = self.save_password(values, self.new_password)
                    if saved:
                        self.window['Save'].update(disabled=True)  # Disable the Save button
                        password_generated = False
                else:
                    sg.popup('No password generated yet!', title='Error')

    def generate_password(self, values):
        if not self.check_inputs():
            return
        # Combine all possible characters into one string
        all_options = string.ascii_letters + string.punctuation + string.digits

        # Generate the initial list of characters
        chars = random.choices(all_options, k=int(values['total_chars']))

        # Ensure that there is at least one uppercase and one lowercase letter
        has_uppercase = any(c.isupper() for c in chars)
        has_lowercase = any(c.islower() for c in chars)
        if not has_uppercase:
            chars[random.randint(0, len(chars) - 1)] = random.choice(string.ascii_uppercase)
        if not has_lowercase:
            chars[random.randint(0, len(chars) - 1)] = random.choice(string.ascii_lowercase)

        # Convert the list of characters to a string and return the password
        new_pass = ''.join(chars)
        return new_pass

    def save_password(self, values, password):
        # Read the existing passwords file or create a new one if it doesn't exist
        try:
            df = pd.read_excel('passwords.xlsx', header=0, skiprows=1)
        except FileNotFoundError:
            # Create a new file if the file doesn't exist
            with open('passwords.xlsx', 'w',):
                pass

            # Create a new empty DataFrame
            df = pd.DataFrame(columns=['Address', 'UserEmail', 'Password'])

        # Check if a password already exists for the given address and user email
        existing_password = df.loc[(df['Address'] == values['address']) & (df['UserEmail'] == values['userEmail']),
                                   'Password'].iloc[0] if len(df.loc[(df['Address'] == values['address']) &
                                                                     (df['UserEmail'] == values[
                                                                         'userEmail'])]) > 0 else None
        if existing_password is not None:
            # A password already exists for the given address and user email
            if existing_password != password:
                # Ask the user if they want to replace the existing password
                response = sg.popup_yes_no(
                    f"A password already exists for the address {values['address']} and user email {values['userEmail']}.\n\nDo you want to replace it with the new password?",
                    title="Replace Password")
                if response == 'Yes':
                    # Replace the existing password with the new one
                    df.loc[(df['Address'] == values['address']) & (
                            df['UserEmail'] == values['userEmail']), 'Password'] = password
                    df.to_excel('passwords.xlsx', index=False, header=True, startrow=1)
                    sg.popup("Password replaced successfully!")
            else:
                sg.popup("The password you entered is already saved for this address and user email.")
        else:
            # Add the new password to the passwords file
            new_password = {'Address': values['address'], 'UserEmail': values['userEmail'], 'Password': password}
            df = df.append(new_password, ignore_index=True)
            df.to_excel('passwords.xlsx', index=False, header=True, startrow=1)
            sg.popup("Password saved successfully!")

    def check_inputs(self):
        # Check if address is not empty
        if not self.window['address'].get():
            sg.popup('Please enter an address!', title='Error')
            return False

        # Check if username/email is not empty and valid
        email = self.window['userEmail'].get()
        if not email:
            sg.popup('Please enter a username or email!', title='Error')
            return False
        elif '@' not in email or '.' not in email:
            sg.popup('Invalid email address', title='Error')
            return False

        return True

    def is_valid_email(email):
        try:
            # Use email.utils.parseaddr to extract the email address
            _, addr = email.utils.parseaddr(email)

            # Use a regular expression to validate the domain name
            domain_regex = re.compile(r'^[a-z0-9]+([.-][a-z0-9]+)*\.[a-z]{2,}$', re.IGNORECASE)
            match = domain_regex.match(addr.split('@')[1])
            if not match:
                return False

            # Use the built-in email validation function to check the full email address
            return bool(email.utils.parseaddr(email)[1])
        except:
            return False


generatePassword = PasswordGenerator()
generatePassword.start()
