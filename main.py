import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
from tkinter import ttk, messagebox
from tkinter import *
from datetime import datetime
import os
import sqlite3

class TiffinServiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tiffin Service Management System")
        
        # Make main window full screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")
        
        # Configure styles
        self.setup_styles()
        
        # File path for storing order history
        self.history_file = "order_history.xlsx"
        
        # Static menu items list instead of Excel
        #self.menu_items = [
        #    {"Item Name": "Dal Rice", "Price": 120},
        #    {"Item Name": "Veg Thali", "Price": 150},
        #    {"Item Name": "Roti Sabzi", "Price": 100},
        #    {"Item Name": "Paneer Special", "Price": 180},
        #    {"Item Name": "South Indian Thali", "Price": 200}
        #]
        
        # Initialize order items and load order history
        self.order_items = []
        self.order_history = self.load_order_history()
        
        # Create main menu
        self.create_menu()
        
        # Create main frames
        self.create_main_frames()
        
    def setup_styles(self):
        # Configure styles for the entire application
        style = ttk.Style()
        
        # Main heading style (for window titles)
        style.configure('Heading.TLabel', 
                       font=('Helvetica', 20, 'bold'),
                       padding=10)
        
        # Subheading style (for section titles)
        style.configure('SubHeading.TLabel', 
                       font=('Helvetica', 16, 'bold'),
                       padding=8)
        
        # Normal text style
        style.configure('Normal.TLabel',
                       font=('Helvetica', 12),
                       padding=5)
        
        # Button styles
        style.configure('Big.TButton',
                       font=('Helvetica', 12, 'bold'),
                       padding=8)
        
        # Treeview styles
        style.configure('Treeview',
                       font=('Helvetica', 11),
                       rowheight=25)
        style.configure('Treeview.Heading',
                       font=('Helvetica', 12, 'bold'))
        
        # Radio button style
        style.configure('TRadiobutton',
                       font=('Helvetica', 12))
        
        # Combobox style
        style.configure('TCombobox',
                       font=('Helvetica', 12))

    def load_order_history(self):
        try:
            orders = []
            # Check if MenuToday.xlsx exists, if not, create it with default structure and sample items
            if not os.path.exists('MenuToday.xlsx'):
                # Create a DataFrame with default columns and sample data
                sample_data = {
                    'Item Name': ['Dal Rice', 'Veg Thali', 'Roti Sabzi', 'Paneer Special', 'South Indian Thali'],
                    'Price': [120, 150, 100, 180, 200]
                }
                default_data = pd.DataFrame(sample_data)
                default_data.to_excel('MenuToday.xlsx', index=False)

            # Load items from MenuToday.xlsx
            items_df = pd.read_excel('MenuToday.xlsx')  # Load the items from the specified file
            items_list = items_df.to_dict(orient='records')  # Convert to a list of dictionaries
            
            # Assuming the items_df has columns 'Item Name' and 'Price'
            self.menu_items = [{'Item Name': item['Item Name'], 'Price': item['Price']} for item in items_list]

            # Load order history from existing order history files
            files = [f for f in os.listdir() if f.startswith('order_history_') and f.endswith('.xlsx')]
            
            for file in files:
                if os.path.exists(file):
                    df = pd.read_excel(file)
                    for _, row in df.iterrows():
                        try:
                            items = eval(str(row['Items']))  # Safely evaluate the items string
                            order = {
                                'order_id': str(row.get('Order ID', '')),
                                'date': str(row['Date']),
                                'items': items,
                                'total': float(row['Total']),
                                'payment_status': str(row.get('Payment Status', 'Not Paid')),
                                'order_type': str(row.get('Order Type', 'Take Away')),
                                'delivery_status': str(row.get('Delivery Status', 'Pending')),
                                'customer_name': str(row.get('Customer Name', ''))
                            }
                            orders.append(order)
                        except Exception as e:
                            print(f"Error processing row: {e}")
                            continue
            return orders
                
        except Exception as e:
            messagebox.showerror("Error", f"Error loading order history: {str(e)}")
    
    def save_order_history(self):
        try:
            # Get all unique dates from order history
            dates = set(order['date'].split()[0] for order in self.order_history)
            
            # For each date, save its orders to corresponding file
            for date in dates:
                filename = f"order_history_{date}.xlsx"
                
                # Filter orders for this date
                date_orders = []
                for order in self.order_history:
                    if order['date'].split()[0] == date:
                        date_orders.append({
                            'Order ID': order['order_id'],
                            'Date': order['date'],
                            'Items': str(order['items']),
                            'Total': order['total'],
                            'Payment Status': order.get('payment_status', 'Not Paid'),
                            'Order Type': order.get('order_type', 'Take Away'),
                            'Delivery Status': order.get('delivery_status', 'Pending'),
                            'Customer Name': order.get('customer_name', '')
                        })
                
                # Save to Excel if we have orders
                if date_orders:
                    df = pd.DataFrame(date_orders)
                    df.to_excel(filename, index=False)

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Error saving order history: {str(e)}")
            return False

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Management Menu
        manage_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Management", menu=manage_menu)
        manage_menu.add_command(label="Customers", command=self.show_customers)
        manage_menu.add_command(label="Menu Items", command=self.show_menu)
        manage_menu.add_command(label="Orders", command=self.show_orders)
        
    def create_main_frames(self):
        # Configure grid weight to enable centering
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Main container frame
        main_container = ttk.Frame(self.root)
        main_container.grid(row=0, column=0)
        
        # Welcome Frame with title and buttons
        self.welcome_frame = ttk.Frame(main_container, padding="40")
        self.welcome_frame.pack(expand=True)
        
        # Title Label with new style
        title_label = ttk.Label(self.welcome_frame, 
                               text="||  अक्षयपात्र  ||",
                               style='Heading.TLabel')
        title_label.pack(pady=40)
        
        # Button Frame
        button_frame = ttk.Frame(self.welcome_frame)
        button_frame.pack(expand=True)
        
        # Quick Actions Buttons with adjusted sizes
        ttk.Button(button_frame, 
                  text="New Order",
                  style='Big.TButton',
                  width=25,  # Adjusted width
                  command=self.new_order).pack(pady=20)
        
        ttk.Button(button_frame,
                  text="View Orders",
                  style='Big.TButton',
                  width=25,
                  command=self.show_orders).pack(pady=20)
        
        ttk.Button(button_frame,
                  text="Manage Customers",
                  style='Big.TButton',
                  width=25,
                  command=self.show_customers).pack(pady=20)
    
    def new_order(self):
        # Create new window for order
        order_window = Toplevel(self.root)
        order_window.title("New Order")
        
        # Make order window full screen
        screen_width = order_window.winfo_screenwidth()
        screen_height = order_window.winfo_screenheight()
        order_window.geometry(f"{screen_width}x{screen_height}+0+0")

        # Configure grid weights for order window
        order_window.grid_columnconfigure(0, weight=1)
        order_window.grid_columnconfigure(1, weight=1)
        order_window.grid_rowconfigure(0, weight=1)

        # Create frames
        item_frame = ttk.Frame(order_window, padding="20")
        item_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=20)
        
        cart_frame = ttk.Frame(order_window, padding="20")
        cart_frame.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.E, tk.W), padx=20)

        # Configure grid weights for frames
        item_frame.grid_columnconfigure(0, weight=1)
        cart_frame.grid_columnconfigure(0, weight=1)

        # Customer Details Section
        ttk.Label(item_frame, text="Customer Name:", style='Normal.TLabel').grid(row=0, column=0, pady=5)
        customer_name_var = tk.StringVar()  # Variable to hold customer name
        customer_name_entry = ttk.Entry(item_frame, textvariable=customer_name_var, width=30)
        customer_name_entry.grid(row=0, column=1, pady=5)

        ttk.Label(item_frame, text="Phone Number:", style='Normal.TLabel').grid(row=1, column=0, pady=5)
        customer_phone_var = tk.StringVar()  # Variable to hold customer phone
        customer_phone_entry = ttk.Entry(item_frame, textvariable=customer_phone_var, width=30)
        customer_phone_entry.grid(row=1, column=1, pady=5)

        ttk.Label(item_frame, text="Address:", style='Normal.TLabel').grid(row=2, column=0, pady=5)
        customer_address_var = tk.StringVar()  # Variable to hold customer address
        customer_address_entry = ttk.Entry(item_frame, textvariable=customer_address_var, width=30)
        customer_address_entry.grid(row=2, column=1, pady=5)

        # Menu Items List with adjusted sizes
        ttk.Label(item_frame, text="Menu Items", style='SubHeading.TLabel').grid(row=3, column=0, columnspan=2, pady=20)
        
        menu_tree = ttk.Treeview(item_frame, columns=('Item', 'Price'), show='headings', height=15, style='Treeview')
        menu_tree.heading('Item', text='Item')
        menu_tree.heading('Price', text='Price')
        menu_tree.column('Item', width=250)
        menu_tree.column('Price', width=100)
        menu_tree.grid(row=4, column=0, columnspan=2, pady=10, sticky=(tk.N, tk.S, tk.E, tk.W))

        # Add buttons frame for menu items
        item_buttons_frame = ttk.Frame(item_frame)
        item_buttons_frame.grid(row=5, column=0, columnspan=2, pady=10)

        # Cart List with adjusted sizes
        ttk.Label(cart_frame, text="Selected Items", style='SubHeading.TLabel').grid(row=0, column=0, pady=20)
        
        cart_tree = ttk.Treeview(cart_frame, columns=('Item', 'Price'), show='headings', height=12, style='Treeview')
        cart_tree.heading('Item', text='Item')
        cart_tree.heading('Price', text='Price')
        cart_tree.column('Item', width=250)
        cart_tree.column('Price', width=100)
        cart_tree.grid(row=1, column=0, pady=10, sticky=(tk.N, tk.S, tk.E, tk.W))

        # Add buttons frame for cart
        cart_buttons_frame = ttk.Frame(cart_frame)
        cart_buttons_frame.grid(row=2, column=0, pady=10)

        # Total Price Label with larger font
        total_price = StringVar(value="Total: ₹0.00")
        ttk.Label(cart_frame, textvariable=total_price, style='SubHeading.TLabel').grid(row=3, column=0, pady=20)

        # Add Order Type Frame
        order_type_frame = ttk.Frame(cart_frame)
        order_type_frame.grid(row=4, column=0, pady=10)
        
        ttk.Label(order_type_frame, text="Order Type:", style='Normal.TLabel').pack(side=tk.LEFT, padx=10)
        
        order_type = tk.StringVar(value="Take Away")
        take_away_radio = ttk.Radiobutton(order_type_frame, text="Take Away", variable=order_type, value="Take Away", style='TRadiobutton')
        take_away_radio.pack(side=tk.LEFT, padx=10)
        
        delivery_radio = ttk.Radiobutton(order_type_frame, text="Delivery", variable=order_type, value="Delivery", style='TRadiobutton')
        delivery_radio.pack(side=tk.LEFT, padx=10)

        # Delivery Status Frame
        delivery_frame = ttk.Frame(cart_frame)
        delivery_frame.grid(row=5, column=0, pady=10)
        
        ttk.Label(delivery_frame, text="Delivery Status:", style='Normal.TLabel').pack(side=tk.LEFT, padx=10)
        
        delivery_status = tk.StringVar(value="Pending")
        delivery_combo = ttk.Combobox(delivery_frame, textvariable=delivery_status, values=["Pending", "Delivered"], style='TCombobox', state="readonly", width=12)
        delivery_combo.pack(side=tk.LEFT, padx=10)

        # Payment Status Frame
        payment_frame = ttk.Frame(cart_frame)
        payment_frame.grid(row=6, column=0, pady=10)
        
        ttk.Label(payment_frame, text="Payment Status:", style='Normal.TLabel').pack(side=tk.LEFT, padx=10)
        
        payment_status = tk.StringVar(value="Not Paid")
        payment_combo = ttk.Combobox(payment_frame, textvariable=payment_status, values=["Paid", "Not Paid"], style='TCombobox', state="readonly", width=12)
        payment_combo.pack(side=tk.LEFT, padx=10)

        def add_to_cart():
            selection = menu_tree.selection()
            if not selection:
                return
            
            item = menu_tree.item(selection[0])['values']
            cart_tree.insert('', 'end', values=item)
            
            # Update total price
            total = sum(float(cart_tree.item(item)['values'][1].replace('₹', '')) for item in cart_tree.get_children())
            total_price.set(f"Total: ₹{total:.2f}")

        def remove_from_cart():
            selection = cart_tree.selection()
            if not selection:
                return
            
            cart_tree.delete(selection[0])
            
            # Update total price
            total = sum(float(cart_tree.item(item)['values'][1].replace('₹', '')) for item in cart_tree.get_children())
            total_price.set(f"Total: ₹{total:.2f}")

        def place_order():
            # Collect items from cart
            ordered_items = []
            total = 0
            for item_id in cart_tree.get_children():
                item = cart_tree.item(item_id)['values']
                ordered_items.append({
                    'Item Name': item[0],
                    'Price': float(item[1].replace('₹', ''))
                })
                total += float(item[1].replace('₹', ''))

            if not cart_tree.get_children():
                messagebox.showwarning("Warning", "Please add items to cart!")
                return  # Stay on the order page

            # Generate unique order ID using timestamp
            order_id = datetime.now().strftime("%Y%m%d%H%M%S")

            # Add to order history
            order = {
                'order_id': order_id,
                'items': ordered_items,
                'total': total,
                'date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'payment_status': payment_status.get(),
                'order_type': order_type.get(),
                'delivery_status': delivery_status.get(),
                'customer_name': customer_name_var.get() or "N/A",  # Use "N/A" if name is empty
                'customer_phone': customer_phone_var.get() or "N/A",  # Use "N/A" if phone is empty
                'customer_address': customer_address_var.get() or "N/A"  # Use "N/A" if address is empty
            }
            self.order_history.append(order)  # Add order to history
            print(f"Order added: {order}")  # Debugging output

            # Save to Excel file
            if not self.save_order_history():
                messagebox.showerror("Error", "Failed to save order history. Please try again.")
                return  # Stay on the order page

            # Show success message
            messagebox.showinfo("Success", 
                                f"Order placed successfully!\n"
                                f"Order ID: {order_id}\n"
                                f"Order Type: {order_type.get()}\n"
                                f"Delivery Status: {delivery_status.get()}\n"
                                f"Total Amount: ₹{total:.2f}\n"
                                f"Payment Status: {payment_status.get()}\n"
                                f"Customer Name: {order['customer_name']}\n"
                                f"Customer Phone: {order['customer_phone']}\n"
                                f"Customer Address: {order['customer_address']}")

            order_window.destroy()  # Close the order window after showing the success message

        # Add buttons
        ttk.Button(item_buttons_frame, text="Add to Cart", style='Big.TButton', width=20, command=add_to_cart).pack(pady=5)
        
        ttk.Button(cart_buttons_frame, text="Remove Item", style='Big.TButton', width=20, command=remove_from_cart).pack(pady=5)
        
        ttk.Button(cart_buttons_frame, text="Place Order", style='Big.TButton', width=20, command=place_order).pack(pady=5)

        # Close button
        ttk.Button(item_buttons_frame, text="Close", style='Big.TButton', width=20, command=order_window.destroy).pack(side=tk.RIGHT, padx=5)

        # Populate menu items
        for item in self.menu_items:
            menu_tree.insert('', 'end', values=(item['Item Name'], f"₹{item['Price']}"))
    
    def show_orders(self):
        # Create new window for order history
        history_window = Toplevel(self.root)
        history_window.title("Order History")
        
        # Make history window full screen
        screen_width = history_window.winfo_screenwidth()
        screen_height = history_window.winfo_screenheight()
        history_window.geometry(f"{screen_width}x{screen_height}+0+0")

        # Load order history at the start
        self.order_history = self.load_order_history()

        # Configure grid weight for centering
        history_window.grid_rowconfigure(0, weight=1)
        history_window.grid_columnconfigure(0, weight=1)

        # Create frame
        frame = ttk.Frame(history_window, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10)

        # Add filter frame
        filter_frame = ttk.Frame(frame)
        filter_frame.grid(row=0, column=0, pady=5, sticky=(tk.W, tk.E))
        
        # Create Treeview and configure it
        tree = ttk.Treeview(frame, 
                            columns=('Order ID', 'Date', 'Items', 'Total', 'Payment', 'Order Type', 'Delivery Status', 'Customer Name'),
                            show='headings',
                            height=18,
                            style='Treeview')
        
        tree.heading('Order ID', text='Order ID')
        tree.heading('Date', text='Date')
        tree.heading('Items', text='Items')
        tree.heading('Total', text='Total')
        tree.heading('Payment', text='Payment Status')
        tree.heading('Order Type', text='Order Type')
        tree.heading('Delivery Status', text='Delivery Status')
        tree.heading('Customer Name', text='Customer Name')
        
        # Adjusted column widths to ensure text is fully visible
        tree.column('Order ID', width=120, minwidth=120)
        tree.column('Date', width=150, minwidth=150)
        tree.column('Items', width=730, minwidth=250)
        tree.column('Total', width=100, minwidth=100)
        tree.column('Payment', width=120, minwidth=120)
        tree.column('Order Type', width=120, minwidth=120)
        tree.column('Delivery Status', width=120, minwidth=120)
        tree.column('Customer Name', width=200, minwidth=200)
        
        tree.grid(row=1, column=0, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure tags for different order states
        tree.tag_configure('paid_and_delivered', background='#90EE90')  # Light green
        tree.tag_configure('paid_only', background='#FFEB3B')  # Yellow
        tree.tag_configure('pending_delivery', background='#FFB6C1')  # Light red

        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=tree.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        tree.configure(yscrollcommand=scrollbar.set)

        def mark_as_paid():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Please select an order to mark as paid", parent=history_window)
                return
            
            item = tree.item(selected_item[0])
            order_id = str(item['values'][0])
            
            # Update order in memory
            updated = False
            for order in self.order_history:
                if str(order.get('order_id', '')) == order_id:
                    order['payment_status'] = 'Paid'
                    updated = True
                    break
            
            if updated:
                if self.save_order_history():
                    update_orders()
                    messagebox.showinfo("Success", "Order marked as paid successfully!", parent=history_window)
                else:
                    messagebox.showerror("Error", "Failed to save payment status update", parent=history_window)
            else:
                messagebox.showerror("Error", f"Order with ID {order_id} not found", parent=history_window)

        def mark_as_delivered():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Please select an order to mark as delivered", parent=history_window)
                return
            
            item = tree.item(selected_item[0])
            order_id = str(item['values'][0])
            
            # Update order in memory
            updated = False
            for order in self.order_history:
                if str(order.get('order_id', '')) == order_id:
                    order['delivery_status'] = 'Delivered'
                    updated = True
                    break
            
            if updated:
                if self.save_order_history():
                    update_orders()
                    messagebox.showinfo("Success", "Order marked as delivered successfully!", parent=history_window)
                else:
                    messagebox.showerror("Error", "Failed to save delivery status update", parent=history_window)
            else:
                messagebox.showerror("Error", f"Order with ID {order_id} not found", parent=history_window)

        # Add buttons frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, pady=20, columnspan=2)

        # Add action buttons
        ttk.Button(button_frame,
                  text="Mark as Paid",
                  style='Big.TButton',
                  width=15,
                  command=mark_as_paid).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame,
                  text="Mark as Delivered",
                  style='Big.TButton',
                  width=15,
                  command=mark_as_delivered).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame,
                  text="Close",
                  style='Big.TButton',
                  width=15,
                  command=history_window.destroy).pack(side=tk.LEFT, padx=10)

        # Date filter
        ttk.Label(filter_frame, 
                 text="Select Date:", 
                 style='Normal.TLabel').pack(side=tk.LEFT, padx=5)
        
        # Get unique dates and add today if not in list
        today = datetime.now().strftime("%Y-%m-%d")
        all_dates = [order['date'].split()[0] for order in self.order_history]
        dates = sorted(list(set(all_dates + [today])), reverse=True)
        
        date_var = tk.StringVar(value=today)  # Set to today by default
        
        date_combo = ttk.Combobox(filter_frame, 
                                 textvariable=date_var, 
                                 values=dates,
                                 style='TCombobox',
                                 state="readonly",
                                 width=15)
        date_combo.pack(side=tk.LEFT, padx=20)

        # View filter
        ttk.Label(filter_frame, 
                 text="View:", 
                 style='Normal.TLabel').pack(side=tk.LEFT, padx=5)
        
        view_var = tk.StringVar(value="All Orders")
        view_combo = ttk.Combobox(filter_frame, 
                                 textvariable=view_var,
                                 values=["All Orders", "Pending Deliveries"],
                                 style='TCombobox',
                                 state="readonly",
                                 width=20)
        view_combo.pack(side=tk.LEFT, padx=20)

        # Payment Status filter
        ttk.Label(filter_frame, 
                 text="Payment Status:", 
                 style='Normal.TLabel').pack(side=tk.LEFT, padx=5)
        
        payment_var = tk.StringVar(value="All")
        payment_combo = ttk.Combobox(filter_frame, 
                                    textvariable=payment_var,
                                    values=["All", "Paid", "Not Paid"],
                                    style='TCombobox',
                                    state="readonly",
                                    width=15)
        payment_combo.pack(side=tk.LEFT, padx=20)

        # Order Type filter
        ttk.Label(filter_frame, 
                 text="Order Type:", 
                 style='Normal.TLabel').pack(side=tk.LEFT, padx=5)
        
        order_type_var = tk.StringVar(value="All")
        order_type_combo = ttk.Combobox(filter_frame, 
                                    textvariable=order_type_var,
                                    values=["All", "Take Away", "Delivery"],
                                    style='TCombobox',
                                    state="readonly",
                                    width=15)
        order_type_combo.pack(side=tk.LEFT, padx=20)

        def update_orders(*args):
            # Clear current items
            for item in tree.get_children():
                tree.delete(item)
            
            selected_date = date_var.get()
            view_type = view_var.get()
            payment_filter = payment_var.get()
            order_type_filter = order_type_var.get()
            
            # Reload order history to get latest data
            self.order_history = self.load_order_history()
            
            # Counter for orders on selected date
            orders_found = 0
            
            # Populate orders for selected date
            for order in self.order_history:
                try:
                    if order['date'].split()[0] == selected_date:
                        # Skip if viewing pending deliveries and order is not pending
                        if view_type == "Pending Deliveries" and order.get('delivery_status') != 'Pending':
                            continue

                        # Skip based on payment filter
                        payment_status = order.get('payment_status', 'Not Paid')
                        if payment_filter != "All" and payment_status != payment_filter:
                            continue
                        
                         # Skip based on order type filter
                        if order_type_filter != "All" and order.get('order_type', 'Take Away') != order_type_filter:
                            continue

                        orders_found += 1
                        
                        # Format items string nicely
                        items_str = ", ".join([f"{item['Item Name']} (₹{item['Price']})" 
                                             for item in order['items']])
                        
                        order_type = order.get('order_type', 'Take Away')
                        delivery_status = order.get('delivery_status', 'Pending')
                        
                        item_id = tree.insert('', 'end', values=(
                            order.get('order_id', ''),
                            order['date'],
                            items_str,
                            f"₹{float(order['total']):.2f}",
                            payment_status,
                            order_type,
                            delivery_status,
                            order.get('customer_name', '')
                        ))
                        
                        # Apply tags for visual status
                        if payment_status == 'Paid':
                            if delivery_status == 'Delivered':
                                tree.item(item_id, tags=('paid_and_delivered',))
                            else:
                                tree.item(item_id, tags=('paid_only',))
                        elif delivery_status == 'Pending':
                            tree.item(item_id, tags=('pending_delivery',))
                except Exception as e:
                    print(f"Error displaying order: {e}")
                    continue
            
            # Show message if no orders found for the selected date and filters
            if orders_found == 0:
                message = f"No orders found for {selected_date}"
                if payment_filter != "All":
                    message += f" with payment status '{payment_filter}'"
                if view_type == "Pending Deliveries":
                    message += " and pending delivery"
                if order_type_filter != "All":
                    message += f" of type '{order_type_filter}'"
                message += "."
                messagebox.showinfo("Orders", message, parent=history_window)

        # Bind update function to combo boxes
        date_combo.bind('<<ComboboxSelected>>', update_orders)
        view_combo.bind('<<ComboboxSelected>>', update_orders)
        payment_combo.bind('<<ComboboxSelected>>', update_orders)
        order_type_combo.bind('<<ComboboxSelected>>', update_orders)  # Bind order type filter

        # Call update_orders directly after window creation
        update_orders()
    
    def show_customers(self):
        # Create new window for customer management
        customer_window = Toplevel(self.root)
        customer_window.title("Manage Customers")
        
        # Make window full screen
        screen_width = customer_window.winfo_screenwidth()
        screen_height = customer_window.winfo_screenheight()
        customer_window.geometry(f"{screen_width}x{screen_height}+0+0")

        # Configure grid weight for centering
        customer_window.grid_rowconfigure(0, weight=1)
        customer_window.grid_columnconfigure(0, weight=1)

        # Create main frame
        frame = ttk.Frame(customer_window, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10)

        # Create Treeview for customers
        tree = ttk.Treeview(frame, 
                            columns=('ID', 'Name', 'Phone', 'Address'),
                            show='headings',
                            height=18,
                            style='Treeview')
        
        tree.heading('ID', text='Customer ID')
        tree.heading('Name', text='Name')
        tree.heading('Phone', text='Phone Number')
        tree.heading('Address', text='Address')
        
        tree.column('ID', width=100, minwidth=100)
        tree.column('Name', width=200, minwidth=200)
        tree.column('Phone', width=150, minwidth=150)
        tree.column('Address', width=400, minwidth=300)
        
        tree.grid(row=0, column=0, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree.configure(yscrollcommand=scrollbar.set)

        # Create frame for customer details entry
        entry_frame = ttk.Frame(frame)
        entry_frame.grid(row=1, column=0, pady=20, sticky=(tk.W, tk.E))

        # Customer details entry fields
        ttk.Label(entry_frame, text="Name:", style='Normal.TLabel').grid(row=0, column=0, padx=5, pady=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(entry_frame, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(entry_frame, text="Phone:", style='Normal.TLabel').grid(row=0, column=2, padx=5, pady=5)
        phone_var = tk.StringVar()
        phone_entry = ttk.Entry(entry_frame, textvariable=phone_var, width=20)
        phone_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(entry_frame, text="Address:", style='Normal.TLabel').grid(row=0, column=4, padx=5, pady=5)
        address_var = tk.StringVar()
        address_entry = ttk.Entry(entry_frame, textvariable=address_var, width=50)
        address_entry.grid(row=0, column=5, padx=5, pady=5)

        def load_customers():
            try:
                if os.path.exists('customer_database.xlsx'):
                    df = pd.read_excel('customer_database.xlsx')
                    customers = []
                    for _, row in df.iterrows():
                        customers.append((
                            str(row['Customer ID']),
                            str(row['Name']),
                            str(row['Phone']),
                            str(row['Address'])
                        ))
                    return customers
                else:
                    # Create new Excel file with headers if it doesn't exist
                    df = pd.DataFrame(columns=['Customer ID', 'Name', 'Phone', 'Address'])
                    df.to_excel('customer_database.xlsx', index=False)
                    return []
            except Exception as e:
                messagebox.showerror("Error", f"Error loading customers: {str(e)}")
                return []

        def save_customers():
            try:
                customers = []
                for item_id in tree.get_children():
                    item = tree.item(item_id)
                    customers.append({
                        'Customer ID': item['values'][0],
                        'Name': item['values'][1],
                        'Phone': item['values'][2],
                        'Address': item['values'][3]
                    })
                
                df = pd.DataFrame(customers)
                df.to_excel('customer_database.xlsx', index=False)
                return True
            except Exception as e:
                messagebox.showerror("Error", f"Error saving customers: {str(e)}")
                return False

        def update_customer_list():
            # Clear current items
            for item in tree.get_children():
                tree.delete(item)
            
            # Load and display customers
            customers = load_customers()
            for customer in customers:
                tree.insert('', 'end', values=customer)

        def add_customer():
            name = name_var.get().strip()
            phone = phone_var.get().strip()
            address = address_var.get().strip()
            
            if not name or not phone or not address:
                messagebox.showwarning("Warning", "Please fill all fields", parent=customer_window)
                return
            
            # Generate customer ID using timestamp
            customer_id = datetime.now().strftime("%Y%m%d%H%M%S")
            
            # Add new customer to tree
            tree.insert('', 'end', values=(customer_id, name, phone, address))
            
            # Save to Excel
            if save_customers():
                # Clear entry fields
                name_var.set('')
                phone_var.set('')
                address_var.set('')
                messagebox.showinfo("Success", "Customer added successfully!", parent=customer_window)

        def edit_customer():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Please select a customer to edit", parent=customer_window)
                return
            
            # Get current values
            item = tree.item(selected_item[0])
            customer_id = item['values'][0]
            
            # Create edit window
            edit_window = Toplevel(customer_window)
            edit_window.title("Edit Customer")
            edit_window.geometry("600x200")
            
            # Variables for edit window
            edit_name_var = tk.StringVar(value=item['values'][1])
            edit_phone_var = tk.StringVar(value=item['values'][2])
            edit_address_var = tk.StringVar(value=item['values'][3])
            
            # Add fields to edit window
            ttk.Label(edit_window, text="Name:", style='Normal.TLabel').grid(row=0, column=0, padx=5, pady=5)
            ttk.Entry(edit_window, textvariable=edit_name_var, width=30).grid(row=0, column=1, padx=5, pady=5)
            
            ttk.Label(edit_window, text="Phone:", style='Normal.TLabel').grid(row=1, column=0, padx=5, pady=5)
            ttk.Entry(edit_window, textvariable=edit_phone_var, width=30).grid(row=1, column=1, padx=5, pady=5)
            
            ttk.Label(edit_window, text="Address:", style='Normal.TLabel').grid(row=2, column=0, padx=5, pady=5)
            ttk.Entry(edit_window, textvariable=edit_address_var, width=50).grid(row=2, column=1, padx=5, pady=5)
            
            def save_edit():
                name = edit_name_var.get().strip()
                phone = edit_phone_var.get().strip()
                address = edit_address_var.get().strip()
                
                if not name or not phone or not address:
                    messagebox.showwarning("Warning", "Please fill all fields", parent=edit_window)
                    return
                
                # Update tree
                tree.item(selected_item[0], values=(customer_id, name, phone, address))
                
                # Save to Excel
                if save_customers():
                    edit_window.destroy()
                    messagebox.showinfo("Success", "Customer details updated successfully!", parent=customer_window)
            
            ttk.Button(edit_window, 
                      text="Save", 
                      style='Big.TButton',
                      command=save_edit).grid(row=3, column=0, columnspan=2, pady=20)

        def delete_customer():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Please select a customer to delete", parent=customer_window)
                return
            
            if not messagebox.askyesno("Confirm", "Are you sure you want to delete this customer?", parent=customer_window):
                return
            
            # Remove from tree
            tree.delete(selected_item[0])
            
            # Save to Excel
            if save_customers():
                messagebox.showinfo("Success", "Customer deleted successfully!", parent=customer_window)

        # Add buttons frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, pady=20)

        # Add buttons
        ttk.Button(button_frame,
                  text="Add Customer",
                  style='Big.TButton',
                  width=15,
                  command=add_customer).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame,
                  text="Edit Customer",
                  style='Big.TButton',
                  width=15,
                  command=edit_customer).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame,
                  text="Delete Customer",
                  style='Big.TButton',
                  width=15,
                  command=delete_customer).pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame,
                  text="Close",
                  style='Big.TButton',
                  width=15,
                  command=customer_window.destroy).pack(side=tk.LEFT, padx=10)

        # Initial load of customers
        update_customer_list()
    
    def show_menu(self):
        # Create new window for menu management
        menu_window = Toplevel(self.root)
        menu_window.title("Menu Management")
        
        # Make menu window full screen
        screen_width = menu_window.winfo_screenwidth()
        screen_height = menu_window.winfo_screenheight()
        menu_window.geometry(f"{screen_width}x{screen_height}+0+0")
        
        messagebox.showinfo("Menu Items", "Menu management will be implemented here")

    def save_customer_details(self, name, phone, address):
        try:
            # Load existing customers or create a new list
            if os.path.exists('customer_database.xlsx'):
                df = pd.read_excel('customer_database.xlsx')
                customers = df.to_dict(orient='records')
            else:
                customers = []

            # Create a new customer entryy
            customer_id = datetime.now().strftime("%Y%m%d%H%M%S")  # Unique ID
            new_customer = {
                'Customer ID': customer_id,
                'Name': name,
                'Phone': phone,
                'Address': address
            }
            customers.append(new_customer)

            # Save back to Excel
            df = pd.DataFrame(customers)
            df.to_excel('customer_database.xlsx', index=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error saving customer details: {str(e)}")
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = TiffinServiceApp(root)
    root.mainloop() 