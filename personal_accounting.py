import sqlite3
import csv
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os
from tkinter import messagebox
from tkcalendar import DateEntry
import calendar
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta


class PersonalAccountingApp(tk.Frame):
    def __init__(self, master):
        super().__init__(master)

        # Connect to database, if not exist, create and populate with predefined values for accounts, categories
        self.conn = sqlite3.connect(config.DATABASE_FILE)
        self.cursor = self.conn.cursor() # Create a cursor to execute SQL commands
        if self.is_database_empty():
            self.create_tables()
            self.populate_tables()
            self.conn.commit()

        # Create the GUI
        self.master = master
        self.master.title('Personal Accounting App')
        self.master.geometry('800x400')
        
        # Create welcome label, description label, and version label
        tk.Label(self.master, text="Welcome to Personal Accounting App!", font=("Arial", 24)).pack(pady=100)
        tk.Label(self.master, text="This app was created by Kim Ngan Nguyen with the help of ChatGPT. To get started, choose from the menu bar.", font=("Arial", 16), wraplength=500).place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        tk.Label(self.master, text=f"© {datetime.now().year} Personal Accounting. All rights reserved.").pack(side=tk.BOTTOM, pady=5)
        tk.Label(self.master, text="Version 0.1.0 (updated 2023-04-23)").pack(side=tk.BOTTOM, pady=5)

        # Create menu bar
        self.create_menu()


    def create_menu(self):
        menu_bar = tk.Menu(self.master)

        # Create file menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.master.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Create accounts menu
        accounts_menu = tk.Menu(menu_bar, tearoff=0)
        accounts_menu.add_command(label="Add Account", command=self.add_account)
        accounts_menu.add_command(label="View Accounts", command=self.view_accounts)
        menu_bar.add_cascade(label="Accounts", menu=accounts_menu)

        # Create transactions menu
        transactions_menu = tk.Menu(menu_bar, tearoff=0)
        transactions_menu.add_command(label="Add Transaction", command=self.add_transaction)
        transactions_menu.add_command(label="View Transactions", command=self.choose_transaction_range)
        menu_bar.add_cascade(label="Transactions", menu=transactions_menu)

        # Create categories menu
        categories_menu = tk.Menu(menu_bar, tearoff=0)
        categories_menu.add_command(label="Add Category", command=self.add_category)
        categories_menu.add_command(label="View Categories", command=self.view_categories)
        menu_bar.add_cascade(label="Categories", menu=categories_menu)

        self.master.config(menu=menu_bar)
    
    
    def add_category(self):
        """
        Creates a new category and adds it to the database.

        Returns:
            _type_: _description_
        """
        return AddCategoryPopup(self)


    def add_account(self):
        """
        Creates a new account and adds it to the database.

        Returns:
            _type_: _description_
        """
        return AddAccountPopup(self)


    def add_transaction(self):
        return AddTransactionPopup(self)


    def view_accounts(self):
        self.cursor.execute("SELECT * FROM accounts")
        rows = self.cursor.fetchall()

        # Create a new window to display the account list
        account_list_window = tk.Toplevel(self.master)
        account_list_window.title("Account List")

        # Create a treeview widget to display the account list
        tree = ttk.Treeview(account_list_window)
        tree.pack(fill='both', expand=True)

        # Define columns for the treeview
        tree['columns'] = ('ID', 'Account', 'Balance')

        # Set column headings
        tree.column('#0', width=0, stretch=tk.NO)
        tree.column('ID', width=50, anchor=tk.CENTER)
        tree.column('Account', width=150, anchor=tk.CENTER)
        tree.column('Balance', width=100, anchor=tk.CENTER)

        tree.heading('ID', text='ID')
        tree.heading('Account', text='Account')
        tree.heading('Balance', text='Balance')

        # Insert account data into the treeview
        for row in rows:
            tree.insert('', 'end', text='', values=(row[0], row[1], row[2]))

        # Add a close button to the window
        close_button = tk.Button(account_list_window, text="Close", command=account_list_window.destroy)
        close_button.pack(pady=10)


    def view_categories(self):
        self.cursor.execute("SELECT * FROM categories")
        rows = self.cursor.fetchall()

        # Create a new window to display the account list
        account_list_window = tk.Toplevel(self.master)
        account_list_window.title("Category List")

        # Create a treeview widget to display the account list
        tree = ttk.Treeview(account_list_window)
        tree.pack(fill='both', expand=True)

        # Define columns for the treeview
        tree['columns'] = ('ID', 'Category')

        # Set column headings
        tree.column('#0', width=0, stretch=tk.NO)
        tree.column('ID', width=50, anchor=tk.CENTER)
        tree.column('Category', width=150, anchor=tk.CENTER)

        tree.heading('ID', text='ID')
        tree.heading('Category', text='Category')

        # Insert account data into the treeview
        for row in rows:
            tree.insert('', 'end', text='', values=(row[0], row[1]))

        # Add a close button to the window
        close_button = tk.Button(account_list_window, text="Close", command=account_list_window.destroy)
        close_button.pack(pady=10)


    def choose_transaction_range(self):
        # Create a new window
        self.window = Toplevel(self.master)
        self.window.title("Choose Transaction Range")

        # Create a frame to hold the date selection widgets
        frame = Frame(self.window)
        frame.pack(pady=10)

        # Create a label and combobox for selecting the date range type
        range_var = StringVar()
        range_combobox = ttk.Combobox(frame, textvariable=range_var, state="readonly", values=["Month", "Year", "Custom"])
        range_combobox.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        # Create a label and combobox for selecting the month (if month is selected as the date range)
        month_label = Label(frame, text="Select month:")
        month_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)

        month_var = StringVar()
        month_combobox = ttk.Combobox(frame, textvariable=month_var, state="disabled", values=[calendar.month_name[i] for i in range(1, 13)])
        month_combobox.grid(row=1, column=1, padx=5, pady=5, sticky=W)

        # Create a label and combobox for selecting the year (if year is selected as the date range)
        year_label = Label(frame, text="Select year:")
        year_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)

        year_var = StringVar()
        year_combobox = ttk.Combobox(frame, textvariable=year_var, state="disabled", values=list(range(2000, 2101)))
        year_combobox.grid(row=2, column=1, padx=5, pady=5, sticky=W)

        # Create labels and entry widgets for selecting the start and end dates (if custom date range is selected)
        start_label = Label(frame, text="Start date (dd/mm/yyyy):")
        start_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)

        start_entry = Entry(frame, state="disabled")
        start_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)

        end_label = Label(frame, text="End date (dd/mm/yyyy):")
        end_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)

        end_entry = Entry(frame, state="disabled")
        end_entry.grid(row=4, column=1, padx=5, pady=5, sticky=W)

        # Function to enable/disable month and year comboboxes based on the range selection
        def on_range_select(*args):
            if range_var.get() == "Month":
                month_combobox.config(state="readonly")
                year_combobox.config(state="readonly")
                start_entry.config(state="disabled")
                end_entry.config(state="disabled")
                start_entry.setvar("")
                end_entry.setvar("")
            elif range_var.get() == "Year":
                year_combobox.config(state="readonly")
                month_combobox.config(state="disabled")
                start_entry.config(state="disabled")
                end_entry.config(state="disabled")
                month_var.set("")
                start_entry.setvar("")
                end_entry.setvar("")
            else:
                month_combobox.config(state="disabled")
                year_combobox.config(state="disabled")
                start_entry.config(state="normal")
                end_entry.config(state="normal")
                month_var.set("")
                year_var.set("")
        
        range_var.trace("w", on_range_select)
    
        # Create a button to display the transactions
        display_button = Button(frame, text="Show", command=lambda: self.display_transactions(self.window, month_var.get(), year_var.get(), start_entry.get(), end_entry.get()))
        display_button.grid(row=5, column=1, padx=5, pady=5, sticky=W)
        

    def display_transactions(self, parent, month=None, year=None, start_date=None, end_date=None):
        # Construct start and end dates based on input
        if month != "" and year != "":
            month_num = datetime.strptime(month, '%B').month
            start_date = datetime(year=int(year), month=month_num, day=1).date()
            end_date = start_date + relativedelta(months=1, days=-1)
        elif year != "":
            start_date = datetime(year=int(year), month=1, day=1).date()
            end_date = datetime(year=int(year), month=12, day=31).date()
        elif start_date == "" or end_date == "":
            messagebox.showerror("Error", "Please provide a valid date range.")
            return
        else:
            start_date = datetime.strptime(start_date, '%d/%m/%Y').date()
            end_date = datetime.strptime(end_date, '%d/%m/%Y').date()

        # Retrieve the transactions within the selected date range from the database
        transactions = self.get_transactions_by_date(start_date, end_date)
        if len(transactions) == 0:
            messagebox.showinfo("No transactions", "There are no transactions for the selected date range.")
            return
        else:
            # Create a new window to display the transactions
            window = tk.Toplevel(parent)
            window.title("Transactions")
            
            # Create a table to display the transactions
            table_frame = ttk.Frame(window)
            table_frame.pack(side=tk.LEFT, padx=10, pady=10)
            
            table_headers = ['ID', 'Date', 'Account', 'Amount', 'Description']
            table = ttk.Treeview(table_frame, columns=table_headers, show='headings')
            
            for header in table_headers:
                table.heading(header, text=header)

            for transaction in transactions:
                table.insert('', 'end', values=transaction)
            table.pack()

            save_button = ttk.Button(window, text="Save to CSV", command=lambda: self.save_to_csv(transactions))
            save_button.pack(side=tk.RIGHT, padx=5, pady=5)
    

    def save_to_csv(self, transactions):
        # Open a file dialog to choose a file name to save the CSV file
        file_name = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not file_name:
            return

        # Write the transactions to the CSV file
        with open(file_name, "w", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["Date", "Account", "Amount", "Category", "Notes"])
            for transaction in transactions:
                writer.writerow(transaction)

        messagebox.showinfo("CSV File Saved", f"The transactions have been saved to {file_name}.")


    def get_transactions_by_date(self, start_date, end_date):
        if start_date.year == end_date.year and start_date.month == end_date.month:
            # Only get transactions for a specific month
            start_of_month = start_date.replace(day=1)
            end_of_month = (end_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
            query = "SELECT * FROM transactions WHERE date BETWEEN ? AND ?"
            self.cursor.execute(query, (start_of_month, end_of_month))
        else:
            # Get transactions for a date range
            query = "SELECT * FROM transactions WHERE date BETWEEN ? AND ?"
            self.cursor.execute(query, (start_date, end_date))

        transactions = self.cursor.fetchall()

        return transactions


    def get_account_list(self):
        """
        Retrieve the list of accounts from the database.

        Returns:
            a list of tuples containing the account names and their balances
        """
        self.cursor.execute("SELECT name FROM accounts")

        return [x[0] for x in self.cursor.fetchall()]


    def get_category_list(self):
        """
        Retrieve the list of categories from the database.

        Returns:
            a list of tuples containing the category names
        """
        self.cursor.execute("SELECT name FROM categories")
        return [x[0] for x in self.cursor.fetchall()]


    def create_tables(self):
        """
        Create the accounts, categories, and transactions tables.
        """
        self.cursor.execute('''CREATE TABLE accounts
            (id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR UNIQUE, balance FLOAT DEFAULT 0)''')
        self.cursor.execute('''CREATE TABLE categories
            (id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR UNIQUE)''')
        self.cursor.execute('''CREATE TABLE transactions
            (id INTEGER PRIMARY KEY AUTOINCREMENT, date DATE, account CHAR, purpose CHAR, amount FLOAT, currency VARCHAR(3), category CHAR)''')
        

    def populate_tables(self):
        """
        Populate the accounts and categories tables with the values in config.py.
        """
        self.cursor.executemany('INSERT INTO accounts (name) VALUES (?)', config.ACCOUNTS)
        self.cursor.executemany('INSERT INTO categories (name) VALUES (?)', config.CATEGORIES)
        

    def is_database_empty(self):
        self.cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table'")
        result = self.cursor.fetchone()[0]

        return result == 0


class AddCategoryPopup:
    def __init__(self, parent):
        self.parent = parent
        self.popup = tk.Toplevel(parent)
        self.popup.title("Add Category")

        # create label and entry widgets for category name
        tk.Label(self.popup, text="Category Name:").grid(row=0, column=0, padx=5, pady=5)
        self.category_entry = tk.Entry(self.popup)
        self.category_entry.grid(row=0, column=1, padx=5, pady=5)

        # create button to add category
        tk.Button(self.popup, text="Add Category", command=self.add_category).grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        tk.Button(self.popup, text="Close", command=self.popup.destroy).grid(row=2, column=0, columnspan=2, padx=5, pady=5)

    def add_category(self):
        category_name = self.category_entry.get()

        # insert category into database
        db_conn = self.parent.conn
        db_cursor = self.parent.cursor
        db_cursor.execute("INSERT INTO categories (name) VALUES (?)", (category_name,))
        db_conn.commit()

        # close popup window
        messagebox.showinfo("Done", "Category added successfully!")
        self.category_entry.delete(0, END)


class AddAccountPopup:
    def __init__(self, parent):
        self.parent = parent
        self.popup = tk.Toplevel(parent)
        self.popup.title("Add Account")

        # create label and entry widgets for account name
        tk.Label(self.popup, text="Account Name:").grid(row=0, column=0, padx=5, pady=5)
        self.account_entry = tk.Entry(self.popup)
        self.account_entry.grid(row=0, column=1, padx=5, pady=5)

        # create label and entry widgets for account name
        tk.Label(self.popup, text="Balance:").grid(row=1, column=0, padx=5, pady=5)
        self.balance_entry = tk.Entry(self.popup)
        self.balance_entry.grid(row=1, column=1, padx=5, pady=5)

        # create button to add account
        tk.Button(self.popup, text="Add Account", command=self.add_account).grid(row=2, column=0, columnspan=2, padx=5, pady=5)
        tk.Button(self.popup, text="Close", command=self.popup.destroy).grid(row=3, column=0, columnspan=2, padx=5, pady=5)


    def add_account(self):
        account_name = self.account_entry.get()
        balance = self.balance_entry.get() if self.balance_entry.get() != "" else 0

        # insert account into database
        db_conn = self.parent.conn
        db_cursor = self.parent.cursor
        db_cursor.execute("INSERT INTO accounts (name, balance) VALUES (?, ?)", (account_name, balance))
        db_conn.commit()

        # close popup window
        messagebox.showinfo("Done", "Account added successfully!")
        self.account_entry.delete(0, END)
        self.balance_entry.delete(0, END)


class AddTransactionPopup:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Add Transaction")
        
        # Retrieve accounts and categories from database
        self.accounts = self.parent.get_account_list()
        self.categories = self.parent.get_category_list()
        self.currencies = ['SGD', 'USD']

        # Create date entry field
        tk.Label(self.window, text="Date:").grid(row=0, column=0, sticky="E")
        self.date_entry = DateEntry(self.window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_entry.grid(row=0, column=1)
        
        # Create table
        self.table = tk.Frame(self.window)
        self.table_headers = ["Account", "Purpose", "Amount", "Currency", "Category"]
        for col, header in enumerate(self.table_headers):
            tk.Label(self.table, text=header, font="bold").grid(row=0, column=col, padx=5, pady=5)
        self.table.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
        self.add_row()
        
        # Create buttons
        tk.Button(self.window, text="Add Row", command=self.add_row).grid(row=2, column=0, padx=5, pady=5)
        tk.Button(self.window, text="Save and Close", command=self.save_and_close).grid(row=2, column=1, padx=5, pady=5)
        tk.Button(self.window, text="Cancel", command=self.window.destroy).grid(row=2, column=2, padx=5, pady=5)
    

    def add_row(self):
        # Get last row number
        row = len(self.table.grid_slaves()) // len(self.table_headers)
        
        # Iterate through each column and create entry widget
        for col in range(len(self.table_headers)):
            if self.table_headers[col] == "Account":
                entry = ttk.Combobox(self.table, values=self.accounts)
            elif self.table_headers[col] == "Category":
                entry = ttk.Combobox(self.table, values=self.categories)
            elif self.table_headers[col] == "Currency":
                entry = ttk.Combobox(self.table, values=self.currencies)
                entry.current(0)
            else:
                entry = tk.Entry(self.table)
            entry.grid(row=row, column=col, padx=5, pady=5)
    

    def save_and_close(self):
        # Get all data from the table and date entry
        date_string = self.date_entry.get()
        date_obj = datetime.strptime(date_string, '%d/%m/%Y').date()

        # Check if date is valid
        if date_string == "":
            messagebox.showerror("Error", "Please enter a valid date.")
            return
        for row in range(1, len(self.table.grid_slaves()) // len(self.table_headers) + 1): # skip header row
            for col in range(len(self.table_headers)):
                entry_list = self.table.grid_slaves(row=row, column=col)
                if entry_list:
                    entry = entry_list[0]
                    if entry.get() == "":
                        messagebox.showerror("Error", "Please enter a valid value for all fields.")
                        return

        # Get data from table
        data = []
        for row in range(1, len(self.table.grid_slaves()) // len(self.table_headers) + 1): # skip header row
            row_data = []
            for col in range(len(self.table_headers)):
                entry_list = self.table.grid_slaves(row=row, column=col)
                if entry_list:  # somehow there is an empty row below header
                    # print(row, col)
                    entry = entry_list[0]
                    row_data.append(entry.get())
            if row_data:
                data.append(row_data)

        # Save transaction
        data_to_save = [[date_obj] + row for row in data]
        self.save_transaction(data_to_save)

        # Update account balance
        account_to_update = {}
        for row in data_to_save:
            account = row[1]
            amount = float(row[3])
            if account in account_to_update:
                account_to_update[account] += amount
            else:
                account_to_update[account] = amount
        self.update_account_balance(account_to_update)

        # Close the window
        messagebox.showinfo("Done", "Transaction added successfully!")
        self.window.destroy()


    def save_transaction(self, data):
        # insert transaction into database
        db_conn = self.parent.conn
        db_cursor = self.parent.cursor
        db_cursor.executemany("INSERT INTO transactions (date, account, purpose, amount, currency, category) VALUES (?, ?, ?, ?, ?, ?)", data)
        db_conn.commit()
        

    def update_account_balance(self, data):
        # update account balance in database
        db_conn = self.parent.conn
        db_cursor = self.parent.cursor
        for account, amount in data.items():
            db_cursor.execute("UPDATE accounts SET balance = balance + ? WHERE name = ?", (amount, account))
        db_conn.commit()


class config:
    DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
    DATABASE_FILE = os.path.join(DATA_DIR, 'expense.db')

    # Default accounts
    ACCOUNTS = [('Bank',), 
                ('Paylah',), 
                ('GrabPay',), 
                ('ShopeePay',), 
                ('Cash',), 
                ('DBS Credit Card',)]
    CATEGORIES = [('Eat out',), ('Groceries',), ('Transport',), ('Others',)]


def run():
    """
    Run the application
    """
    # Create the data directory if it doesn't exist
    os.makedirs(os.path.join(os.path.dirname(__file__), 'data'), exist_ok=True)
        
    # create GUI window
    root = Tk()
    root.title('Personal Accounting App')
    app = PersonalAccountingApp(root)
    root.mainloop()


if __name__ == '__main__':
    run()
