import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from tkinter import filedialog, ttk, messagebox, simpledialog
import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import seaborn as sns
import warnings
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar

warnings.filterwarnings("ignore", category=RuntimeWarning)

class BusinessManagementSystem:
    def __init__(self, root):
        self.root = root
        self.df = None
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Business Management System ┃ Developed by Grp-No.-102")
        self.root.geometry("1200x1000")
        self.root.configure(bg="#05113b")
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # setting up custom styles for each selected widgets
        self.setup_custom_styles()

        self.setup_header()

        
        self.style.configure('TNotebook',
                        background='#05113b',       # Background of the tab area
                        borderwidth=2,        # Border width
                        tabmargins=[0,0,0,0],   # Margins around the tabs
                            )
        self.style.configure('TNotebook.Tab',anchor="center",
                            font=('Callibri', 18, 'bold'),  # Tab font
                            padding=[10,10],             # Padding inside tab
                            width=40,
                            focuscolor='None',   # Border color when focused
                            )

        # Create main notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both')

        # Create tabs
        self.setup_data_tab()
        self.setup_sales_tab()
        self.setup_inventory_tab()
        self.setup_finance_tab()
        self.setup_reporting_tab()
        self.setup_market_trends_tab()
        root.update_idletasks()

       

        # Status bar at bottom
        self.status_var = tk.StringVar()
        self.status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.FLAT, anchor=tk.W,bg="#05113b",font=("Calibri",12))
        
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X,ipadx=45,ipady=5)
        self.status_var.set("Ready")
    
    
    def setup_header(self):
         
        
        self.frame_header=tk.Frame(root,bg="#05113b",height=110)      # height is important for the header frame

        self.frame_header.pack(side="top",expand=False,fill="both")
        self.frame_header.pack_propagate(False)

        # ~~~~~~~~ logo ~~~~~~~~~

        # === QUICK IMAGE LOAD & RESIZE === using pillow
        # currently not applying enchancer and 
        img = Image.open("logo.png")                     # Changing this to logo for the app
        img = img.resize((450,90), Image.LANCZOS)        # Resize(adjusting as needed)
        
        self.logo=ImageTk.PhotoImage(img)
        self.label_logo=tk.Label(self.frame_header,image=self.logo,compound="center",bg="#05113b")
        self.label_logo.place(height=100,relx=0,rely=0.0) 

        # ~~~~~~~~~ exit button ~~~~~~~~~~~
        self.but_exit = ttk.Button(self.frame_header, text="Exit", underline=0, cursor="hand2", 
        style="Exit.TButton", command=self.exit_application)  # Added command
        self.but_exit.place(relx=0.9, rely=0.1, relwidth=0.1, relheight=0.4) 

        # ~~~~~~~~~ about us button ~~~~~~~~~~
        self.but_about_us = ttk.Button(self.frame_header, text="About us", underline=0, cursor="hand2", 
        style="Aboutus.TButton", command=self.show_about_us)  # Added command
        self.but_about_us.place(relx=0.9, rely=0.5, relwidth=0.1, relheight=0.4)  
        
    #function to setup custom styles for types of widgets
    def setup_custom_styles(self):
        # Create custom style
        self.style = ttk.Style()

        # Choose a theme (use 'clam' if default theme ignores colors)
        self.style.theme_use('clam')

        # styling for button EXIT
        self.style.configure("Exit.TButton",bg="#D97272",
                             focusthickness=0,focuscolor='none',
                             font=("Calibri",12),relief="flat"
                            )
        self.style.map("Exit.TButton",
                       foreground=[
                                        ("active", "#fff703"),     # Hover color
                                        ("pressed", "#D72020"),    # When clicked
                                        ("selected", "#5324d2"),   # If button is "selected"
                                        ("!active", "#D97272")     # Default
                                    ],
                       background=[
                                        ("active", "#d97272"),
                                        ("pressed", "white"),
                                        ("!active", "#05113b")
                                    ]
                     )
        
        # styling for button About Us
        self.style.configure("Aboutus.TButton",focuscolor="none",focusthickness=0,
                             font=("Calibri",12),relief="flat")
        self.style.map("Aboutus.TButton",
                       foreground=[
                                        ("active", "#32b7a5"),     # Hover color
                                        ("pressed", "#66aea4"),    # When clicked
                                        ("selected", "#5324d2"),   # If button is "selected"
                                        ("!active", "white")     # Default
                                    ],
                       background=[
                                        ("active", "#05113b"),
                                        ("pressed", "white"),
                                        ("!active", "#05113b")
                                    ]
                     )

        # ~~~~~~~~~~~ custom style for TNotebook and its tabs ~~~~~~~~~~

        self.style.configure('TNotebook',
                        background='#05113b',       # Background of the tab area
                        borderwidth=2,        # Border width
                        tabmargins=[0,0,0,0],   # Margins around the tabs
                            )
        # for each of the tabs
        self.style.configure('TNotebook.Tab',anchor="center",
                            font=('Callibri', 18, 'bold'),  # Tab font
                            padding=[10,10],             # Padding inside tab
                            width=40,
                            focuscolor='None',   # Border color when focused
                            )
        self.style.map('TNotebook.Tab',
                        background=[("active","#453F3F"),     # give active before selected,!selected
                                    ('selected', "#141414"),
                                     ('!selected', '#1d48e1'),
                                     
                        ],
                        foreground=[('selected', 'white'), ('!selected', 'white')],
                        padding=[
                                ('selected', [14, 12]),  # Wider and taller
                                ('!selected', [10, 5])
                        ]
                    )
        # Removing default focus ring (reason:focus rings makes ux worst)
        self.style.layout("TNotebook.Tab", [
            ('Notebook.tab', {
                'sticky': 'nswe',
                'children': [
                    ('Notebook.padding', {
                        'side': 'top',
                        'sticky': 'nswe',
                        'children': [
                            ('Notebook.label', {'side': 'left', 'sticky': ''})
                        ]
                    })
                ]
            })
        ])


        # style for success buttons
        self.style.configure("Success.TButton",
                    background="#4CAF50",foreground="black",cursor="hand2",border=0,relief="flat",width=20,
                    focusthickness=0,font=("Calibri",12),
                    focuscolor='none')

        # map style for style buttons
        self.style.map("Success.TButton",
                    background=[
                    ("active", "#377A3A"),     # Hover color
                    ("pressed", "#377A3A"),    # When clicked
                    ("selected", "#5324d2"),   # If button is "selected"
                    ("!active", "#4CAF50")     # Default
                    ],
                    foreground=[
                        ("disabled", "gray"),
                        ("pressed", "#b2abab"),
                        ("active", "#fff703"),
                        ("!active","white")
                    ]
                )
        
        # styling for export button
        self.style.configure("Export.TButton",
                    background="#4CAF50",foreground="black",border=0,relief="flat",width=20,
                    focusthickness=0,font=("Calibri",12),
                    focuscolor='none')

        # map style for exprt button
        self.style.map("Export.TButton",
                    background=[
                    ("active", "#026381"),     # Hover color
                    ("pressed", "#354EDC"),    # When clicked
                    ("selected", "#121315"),   # If button is "selected"
                    ("!active", "#0091BD")     # Default
                    ],
                    foreground=[
                        ("disabled", "gray"),
                        ("pressed", "#b2abab"),
                        ("active", "#fff703"), ("!active","white")
                    ]
                )
        
        # style for the label frame containing buttons
        self.style.configure("Buttons.TLabelframe",
            background="#141414",
            bordercolor="#0c173f",     # Blue border
            borderwidth=30,
            relief="solid"
        )

        # for the label of TLabelframe
        self.style.configure("Buttons.TLabelframe.Label",
            background="#141414",
            foreground="#1d48e1",
            font=("Calibri", 12, "bold"),
            padding=4
        )

        # style for tool  buttons
        self.style.configure("ToolButtons.TButton",
                    background="#4CAF50",foreground="black",border=0,relief="flat",
                    focusthickness=0,
                    focuscolor='none',padding=(10,12))

        # map style for tool buttons
        self.style.map("ToolButtons.TButton",
                    background=[
                    ("active", "#3751E8"),     # Hover color
                    ("pressed", "#354EDC"),    # When clicked
                    ("selected", "#121315"),   # If button is "selected"
                    ("!active", "#585C75")     # Default
                    ],
                    foreground=[
                        ("disabled", "gray"),
                        ("pressed", "#274370"),
                        ("active", "#fff703"),
                        ("!active","white")
                    ]
        )

        self.style.configure("Custom.TCombobox",
                        fieldbackground="green",     # background inside the entry field
                        background="#575353",        # dropdown arrow background
                        foreground="black",       # text color
                        bordercolor="#575353",

                        lightcolor="#575353",
                        darkcolor="#575353",
                        relief="flat",
                        arrowcolor="white",
                        padding=4,

        )
        
        # Custom scrollbar style
        self.style.configure("Custom.Vertical.TScrollbar",
                        gripcount=0,
                        background="#274370",
                        troughcolor="#1e1e1e",
                        bordercolor="#05d146",
                        arrowcolor="#05d146",
                        relief="flat"
        )

        #  style for LabelFrame of Result
        self.style.configure("Result.TLabelframe",
            background="#141414",
            bordercolor="#0c173f",     # Blue border
            borderwidth=5,
            relief="solid"
        )

        # style for the label of the labelframe of the result
        self.style.configure("Result.TLabelframe.Label",
            background="#141414",
            foreground="#1d48e1",
            font=("Calibri", 12, "bold"),
            padding=4
        )


    def setup_data_tab(self):
        """Tab for data loading and basic analysis"""
        data_tab = tk.Frame(self.notebook,bg="#141414")
        self.notebook.add(data_tab, text="Data Management")
        
        # Header
        header = tk.Label(data_tab, text="Load your Excel file to get started",anchor="e", font=("Calibri", 13, "bold"), bg="#141414", fg="white")
        header.pack(fill="x",pady=10,padx=10)
        
        # File loading frame
        load_frame = tk.Frame(data_tab, bg="#141414", bd=1, relief="flat")
        load_frame.pack(padx=10, pady=10, fill="both")
        
        
        load_btn = ttk.Button(load_frame,style="Success.TButton", text="Load Excel File ⬇️",cursor="hand2",command=self.load_file)
        load_btn.pack(pady=10, padx=10, side="left")


        self.export_btn = ttk.Button(load_frame, style="Export.TButton",text="Export Data ⬆️",cursor="hand2",
                                   state=tk.DISABLED, command=self.export_data)
        self.export_btn.pack(pady=10, padx=10, side="right")
        
        

        

        # Data cleaning frame
        clean_frame = ttk.LabelFrame(data_tab, text="Data Cleaning Tools",style="Buttons.TLabelframe")
        clean_frame.pack(pady=10, padx=10, fill="x")
        
         
         
        buttons = [
            # ("Fill Missing with Mean", "#607D8B", self.fill_missing_with_mean),
            ("Fill Missing with Mean", "ToolButtons.TButton", self.fill_missing_with_mean),

            ("Fill Missing with 0", "ToolButtons.TButton", self.fill_missing_with_zero),
            ("Drop Missing Rows", "ToolButtons.TButton", self.drop_missing_rows),
            ("Drop Duplicates", "ToolButtons.TButton", self.drop_duplicates),
            ("Reset Index", "ToolButtons.TButton", self.reset_index),
            ("Show Missing Data", "ToolButtons.TButton", self.show_missing_summary)
        ]
        
        for text, style_buttons, command in buttons:
            btn = ttk.Button(clean_frame, text=text,style=style_buttons,cursor="hand2", command=command)
            btn.pack(side="left", padx=2, pady=2, expand=True,fill="x")
        
        # Basic analysis frame
        analysis_frame = ttk.LabelFrame(data_tab, text="Basic Analysis",style="Buttons.TLabelframe")
        analysis_frame.pack(pady=10, padx=10, fill="x")
        
        # Column selection
        self.column_var = tk.StringVar()
        col_select_frame = tk.Frame(analysis_frame, bg="#141414")
        col_select_frame.pack(pady=5, fill="x")
        
        tk.Label(col_select_frame, text="Select Column:", bg="#141414",fg="white").pack(side="left", padx=5)
        self.column_dropdown = ttk.Combobox(col_select_frame,style="Custom.TCombobox", textvariable=self.column_var, state="readonly")
        self.column_dropdown.pack(side="left", padx=5, fill="x", expand=True)
        
        

        # Analysis buttons
        btn_frame = tk.Frame(analysis_frame, bg="#141414")
        btn_frame.pack(pady=5)
        
        analysis_buttons = [
            ("Calculate Stats", "ToolButtons.TButton", self.calculate_stats),
            ("Histogram", "ToolButtons.TButton", self.plot_histogram),
            ("Boxplot", "ToolButtons.TButton", self.plot_boxplot),
            ("Filter Data", "ToolButtons.TButton", self.show_filter_dialog)
        ]
        
        for text, style_buttons, command in analysis_buttons:
            btn = ttk.Button(btn_frame, text=text, style=style_buttons,cursor="hand2", command=command)
            btn.pack(side="left",padx=2)
        
        # Results display
        result_frame = tk.Frame(data_tab, bg="#141414", bd=0, relief="flat")
        result_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.result_box = tk.Text(result_frame, wrap="word",borderwidth=0,highlightthickness=0, font=("Consolas", 13), bg="#1e1e1e", fg="white",state=tk.DISABLED)
        
        self.result_box.pack(fill="both", expand=True)                                           
        
        # adding scrollbar to the result box Text
        scrollbar = ttk.Scrollbar(self.result_box,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.result_box.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.result_box.yview)
        
    
    def setup_sales_tab(self):
        """Tab for sales analysis"""
        sales_tab = tk.Frame(self.notebook,bg="#141414")
        self.notebook.add(sales_tab, text="Sales Analysis")
        
        # Header
        header = tk.Label(sales_tab, text="Select the Time period to Begin",anchor="e",  font=("Calibri", 13, "bold"), bg="#141414", fg="white")
        header.pack(fill="x",pady=10,padx=10)

        
        # frame containing time and buttons
        frame_time_and_buttons=tk.Frame(sales_tab,bg="#141414")
        frame_time_and_buttons.pack(fill="x",expand=False,padx=10,pady=30)
        # Time period selection
        time_period_frame = tk.Frame(frame_time_and_buttons, bg="#141414")
        time_period_frame.pack(side="left",fill="both",expand=True,pady=10)
        
        tk.Label(time_period_frame, text="Time Period:", bg="#141414").pack(side="left",anchor="ne",expand=True)
        self.period_var = tk.StringVar(value="Monthly")
        period_options = ["Daily", "Weekly", "Monthly", "Quarterly", "Yearly"]
        period_menu = ttk.Combobox(time_period_frame, textvariable=self.period_var,style="Custom.TCombobox", values=period_options, state="readonly")
        period_menu.pack(side="left",expand=True,anchor="nw")
        

        
        # Sales analysis buttons
        btn_frame = tk.Frame(frame_time_and_buttons, bg="#141414")
        btn_frame.pack(side="right",fill="both",expand=True,anchor="n")
        
        sales_buttons = [
            ("Sales Trends", "ToolButtons.TButton", self.show_sales_trends),
            ("Product Performance", "ToolButtons.TButton", self.show_product_performance),
            ("Regional Sales", "ToolButtons.TButton", self.show_regional_sales),
            ("Revenue Distribution", "ToolButtons.TButton", self.show_revenue_distribution),
            ("Peak Hours", "ToolButtons.TButton", self.show_peak_hours),
            ("Returns Analysis", "ToolButtons.TButton", self.show_returns_analysis)
        ]
        
        # Configure 3 columns and 2 rows to have equal size
        for col in range(3):
            btn_frame.columnconfigure(col, weight=1, uniform="equal")
        for row in range(2):
            btn_frame.rowconfigure(row, weight=1, uniform="equal")
        # Arrange buttons in 2x3 grid
        for i, (text, style_buttons, command) in enumerate(sales_buttons):
            btn = ttk.Button(btn_frame, text=text, style=style_buttons,cursor="hand2", command=command)
            btn.grid(row=i//3,sticky="nsew", column=i%3, padx=5, pady=5)
        

        

        # i want the text widget to have a minumum fixed height(number of rows) ,  so creating a void frame
        frame_void=tk.Frame(sales_tab,bg="#141414",height=3)
        frame_void.pack(padx=5,pady=5,fill="both",expand=True)
        # Sales metrics frame
        metrics_frame = ttk.LabelFrame(frame_void, text="Key Sales Metrics",style="Result.TLabelframe")
        metrics_frame.pack(expand=True,fill="both")
        
        self.sales_metrics_text = tk.Text(metrics_frame, wrap="word", font=("Consolas", 13), height=5,borderwidth=0,highlightthickness=0,
                                          bg="#1e1e1e", fg="white")
        self.sales_metrics_text.pack(fill="both", expand=True)
        self.sales_metrics_text.config(state="disabled")

        # adding scrollbar to sales metrics text Text
        scrollbar = ttk.Scrollbar(self.sales_metrics_text,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.sales_metrics_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.sales_metrics_text.yview)
        
        # Chart frame
        self.sales_chart_frame = tk.Frame(sales_tab, bg="#1e1e1e", bd=1, relief="flat")
        self.sales_chart_frame.pack(padx=10, pady=10, fill="both", expand=True)
    
    def setup_inventory_tab(self):
        """Tab for inventory analysis"""
        inventory_tab = tk.Frame(self.notebook,bg="#141414")
        self.notebook.add(inventory_tab, text="Inventory Analysis")
        
        # Header
        header = tk.Label(inventory_tab, text="Select a tool to analyze your inventory performance",anchor="e", font=("Calibri", 13, "bold"), bg="#141414", fg="white")
        header.pack(fill="x",padx=10,pady=10)
        
        # Inventory analysis buttons
        btn_frame = tk.Frame(inventory_tab, bg="#141414")
        btn_frame.pack(pady=20,padx=10,fill="x")
        
        inventory_buttons = [
            ("Reorder Points", "ToolButtons.TButton", self.calculate_reorder_points),
            ("Stock Aging", "ToolButtons.TButton", self.show_stock_aging),
            ("Over/Under Stock", "ToolButtons.TButton", self.show_stock_levels),
            ("Inventory Turnover", "ToolButtons.TButton", self.calculate_inventory_turnover)
        ]

        # Configure 3 columns and 2 rows to have equal size
        for col in range(3):
            btn_frame.columnconfigure(col, weight=1, uniform="equal")
        for row in range(2):
            btn_frame.rowconfigure(row, weight=1, uniform="equal")
        # Arrange buttons in 2x3 grid
        for i, (text, style_buttons, command) in enumerate(inventory_buttons):
            btn = ttk.Button(btn_frame, text=text, style=style_buttons, cursor="hand2",command=command)
            btn.grid(row=i//3,sticky="nsew", column=i%3, padx=5, pady=5)
        
        # i want the text widget to have a minumum fixed height(number of rows) ,  so creating a void frame
        frame_void=tk.Frame(inventory_tab,bg="#141414",height=8)
        frame_void.pack(padx=5,pady=5,fill="both",expand=True)
        # Inventory metrics frame
        metrics_frame = ttk.LabelFrame(frame_void, text="Inventory Metrics", style="Result.TLabelframe")
        metrics_frame.pack(expand=True,fill="both")
        
        self.inventory_metrics_text = tk.Text(metrics_frame, wrap="word", font=("Consolas", 13), borderwidth=0,highlightthickness=0,
                                            height=8, bg="#1e1e1e", fg="white")
        self.inventory_metrics_text.pack(fill="both", expand=True)
        self.inventory_metrics_text.config(state="disabled")

        # adding scrollbar to inventoy metrics text
        scrollbar = ttk.Scrollbar(self.inventory_metrics_text,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.sales_metrics_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.inventory_metrics_text.yview)
        
        # Chart frame
        self.inventory_chart_frame = tk.Frame(inventory_tab, bg="#1e1e1e", bd=1, relief="flat")
        self.inventory_chart_frame.pack(padx=10, pady=10, fill="both", expand=True)
    
    def setup_finance_tab(self):
        """Tab for financial analysis"""
        finance_tab = tk.Frame(self.notebook,bg="#141414")
        self.notebook.add(finance_tab, text="Financial Analysis")
        
        # Header
        header = tk.Label(finance_tab, text="Select a Comparison Interval to begin",anchor="e", font=("Calibri", 13, "bold"), bg="#141414", fg="white")
        header.pack(fill="x",padx=10,pady=10)
        
        # frame containing time and buttons
        frame_time_and_buttons=tk.Frame(finance_tab,bg="#141414")
        frame_time_and_buttons.pack(fill="x",expand=False,padx=10,pady=30)
        # Time period selection
        time_period_frame = tk.Frame(frame_time_and_buttons, bg="#141414")
        time_period_frame.pack(side="left",fill="both",expand=True,pady=0)
        
        tk.Label(time_period_frame, text="Compare:", bg="#141414").pack(side="left",anchor="ne",expand=True)
        self.finance_period_var = tk.StringVar(value="Month-over-Month")
        period_options = ["Month-over-Month", "Year-over-Year", "Quarter-over-Quarter"]
        period_menu = ttk.Combobox(time_period_frame, textvariable=self.finance_period_var, values=period_options,style="Custom.TCombobox", state="readonly")
        period_menu.pack(side="left",expand=True,anchor="nw")
        
        # Finance analysis buttons
        btn_frame = tk.Frame(frame_time_and_buttons, bg="#141414")
        btn_frame.pack(side="right",fill="both",expand=True,anchor="n")
        
        finance_buttons = [
            ("Profit Analysis", "ToolButtons.TButton", self.show_profit_analysis),
            ("Growth Tracking", "ToolButtons.TButton", self.show_growth_tracking),
            ("Financial Ratios", "ToolButtons.TButton", self.show_financial_ratios),
            ("Expense Analysis", "ToolButtons.TButton", self.show_expense_analysis)
        ]
        # Configure 3 columns and 2 rows to have equal size
        for col in range(3):
            btn_frame.columnconfigure(col, weight=1, uniform="equal")
        for row in range(2):
            btn_frame.rowconfigure(row, weight=1, uniform="equal")
        # Arrange buttons in 2x3 grid
        for i, (text, style_buttons, command) in enumerate(finance_buttons):
            btn = ttk.Button(btn_frame, text=text, style=style_buttons, cursor="hand2",command=command)
            btn.grid(row=i//3,sticky="nsew", column=i%3, padx=5, pady=5)
        
        
        # i want the text widget to have a minumum fixed height(number of rows) ,  so creating a void frame
        frame_void=tk.Frame(finance_tab,bg="#141414",height=3)
        frame_void.pack(padx=5,fill="both",expand=True)
        # Finance metrics frame
        metrics_frame = ttk.LabelFrame(frame_void, text="Financial Metrics",style="Result.TLabelframe")
        metrics_frame.pack(expand=True,fill="both")
        
        self.finance_metrics_text = tk.Text(metrics_frame, wrap="word", font=("Consolas", 13), borderwidth=0,highlightthickness=0,
                                          height=8, bg="#1e1e1e", fg="white")
        self.finance_metrics_text.pack(fill="both", expand=True)
        self.finance_metrics_text.config(state="disabled")
        
        # adding scrollbar to metrics text
        scrollbar = ttk.Scrollbar(self.finance_metrics_text,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.sales_metrics_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.finance_metrics_text.yview)

        # Chart frame
        self.finance_chart_frame = tk.Frame(finance_tab, bg="#1e1e1e", bd=1, relief="flat")
        self.finance_chart_frame.pack(padx=10, pady=10, fill="both", expand=True)
    
    def setup_reporting_tab(self):
        """Tab for report generation"""
        reporting_tab = tk.Frame(self.notebook,bg="#141414")
        self.notebook.add(reporting_tab, text="Reporting")
        
        # Header
        header = tk.Label(reporting_tab, text="Generate and Export Detailed Reports",anchor="e", font=("Calibri", 13, "bold"), bg="#141414", fg="white")
        header.pack(fill="x",padx=10,pady=10)
        
        # frame contaning report frame and generate button
        frame_reportType_and_generateButton=tk.Frame(reporting_tab,bg="#141414")
        frame_reportType_and_generateButton.pack(fill="x",pady=10)

        # Report type selection
        report_type_selection_frame = tk.Frame(frame_reportType_and_generateButton, bg="#141414")
        report_type_selection_frame.pack(side="left",padx=5,pady=10)
        
        tk.Label(report_type_selection_frame, text="Report Type:", bg="#141414").pack(side="left", padx=5)
        self.report_type_var = tk.StringVar(value="Sales Summary")
        report_options = ["Sales Summary", "Inventory Status", "Financial Summary", "Comprehensive"]
        report_menu = ttk.Combobox(report_type_selection_frame, textvariable=self.report_type_var,style="Custom.TCombobox", values=report_options, state="readonly")
        report_menu.pack(side="left", padx=5)
        
        # Time period selection
        time_frame = tk.Frame(frame_reportType_and_generateButton, bg="#141414")
        time_frame.pack(side="left",padx=10,pady=10)
        
        tk.Label(time_frame, text="Time Period:", bg="#141414").pack(side="left", padx=5)
        self.report_period_var = tk.StringVar(value="Monthly")
        period_options = ["Daily", "Weekly", "Monthly", "Quarterly", "Yearly"]
        period_menu = ttk.Combobox(time_frame, textvariable=self.report_period_var,style="Custom.TCombobox", values=period_options, state="readonly")
        period_menu.pack(side="left", padx=5)
        
        # Generate report button
        gen_btn = ttk.Button(frame_reportType_and_generateButton, text="Generate Report 📄", style="Success.TButton",cursor="hand2", command=self.generate_report)
        gen_btn.pack(side="left",padx=60,fill="y",pady=5)
        
        # Report preview frame
        preview_frame = ttk.LabelFrame(reporting_tab, text="Report Preview",style="Result.TLabelframe")
        preview_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.report_text = tk.Text(preview_frame, wrap="word", font=("Consolas", 13), borderwidth=0,highlightthickness=0,
                                 height=15, bg="#1e1e1e", fg="white")
        self.report_text.pack(fill="both", expand=True)
        self.report_text.config(state="disabled")
        
        # adding scrollbar to report text
        scrollbar = ttk.Scrollbar(self.report_text,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.sales_metrics_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.report_text.yview)

        # Export buttons
        export_frame = tk.Frame(reporting_tab, bg="#141414")
        export_frame.pack(side="right",pady=10)
        
        ttk.Button(export_frame, text="Export to Excel ⬆️",style="Export.TButton",cursor="hand2", command=self.export_report).pack(side="left", padx=5)
        ttk.Button(export_frame, text="Export to PDF ⬆️", style="Export.TButton",cursor="hand2", command=self.export_pdf).pack(side="left", padx=5)
    
    def setup_market_trends_tab(self):
        """Tab for market trend analysis"""
        trends_tab = tk.Frame(self.notebook,bg="#141414")
        self.notebook.add(trends_tab, text="Market Trends")
        
        # Header
        header = tk.Label(trends_tab, text="Select a Tool to Analyze Market Trends and Performance", font=("Calibri", 13, "bold"),anchor="e", bg="#141414", fg="white")
        header.pack(fill="x",padx=10,pady=10)
        
        # Market analysis buttons
        btn_frame = tk.Frame(trends_tab, bg="#141414")
        btn_frame.pack(padx=10,pady=20,fill="x")
        
        trend_buttons = [
            ("Market Trends", "ToolButtons.TButton", self.analyze_market_trends),
            ("Seasonality", "ToolButtons.TButton", self.analyze_seasonality),
            ("Demand Forecasting", "ToolButtons.TButton", self.forecast_demand),
            ("Competitive Analysis", "ToolButtons.TButton", self.competitive_analysis)
        ]
        
        # Configure 3 columns and 2 rows to have equal size
        for col in range(3):
            btn_frame.columnconfigure(col, weight=1, uniform="equal")
        for row in range(2):
            btn_frame.rowconfigure(row, weight=1, uniform="equal")
        # Arrange buttons in 2x3 grid
        for i, (text, style_buttons, command) in enumerate(trend_buttons):
            btn = ttk.Button(btn_frame, text=text, style=style_buttons,cursor="hand2", command=command)
            btn.grid(row=i//3,sticky="nsew", column=i%3, padx=5, pady=5)
        
        # i want the text widget to have a minumum fixed height(number of rows) ,  so creating a void frame
        frame_void=tk.Frame(trends_tab,bg="#141414",height=8)
        frame_void.pack(padx=5,pady=5,fill="both",expand=True)
        # Analysis results frame
        analysis_frame = ttk.LabelFrame(frame_void, text="Analysis Results", style="Result.TLabelframe")
        analysis_frame.pack(expand=True, fill="both",pady=10 )
        
        self.trends_text = tk.Text(analysis_frame, wrap="word", font=("Consolas", 13), borderwidth=0,highlightthickness=0,
                                 height=8, bg="#1e1e1e", fg="white")
        self.trends_text.pack(fill="both", expand=True)
        self.trends_text.config(state="disabled")
        
        # adding scrollbar to trends text
        scrollbar = ttk.Scrollbar(self.trends_text,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.sales_metrics_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.trends_text.yview)
        

        # Chart frame
        self.trends_chart_frame = tk.Frame(trends_tab, bg="#1e1e1e", bd=1, relief="flat")
        self.trends_chart_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        # ==================== MARKET TRENDS METHODS ====================
    
    def analyze_market_trends(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
            
        if 'Date' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Date' and 'Sales' not found.")
            return
            
        # Prepare data
        df_trends = self.df[['Date', 'Sales']].copy()
        df_trends.set_index('Date', inplace=True)
        monthly_sales = df_trends.resample('M').sum()
        
        # Calculate moving average
        monthly_sales['MA_3'] = monthly_sales['Sales'].rolling(window=3).mean()
        monthly_sales['MA_6'] = monthly_sales['Sales'].rolling(window=6).mean()
        
        # Plot
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(monthly_sales.index, monthly_sales['Sales'], label='Actual Sales', marker='o')
        ax.plot(monthly_sales.index, monthly_sales['MA_3'], label='3-Month Moving Avg', linestyle='--')
        ax.plot(monthly_sales.index, monthly_sales['MA_6'], label='6-Month Moving Avg', linestyle='-.')
        ax.set_title("Market Sales Trends")
        ax.set_xlabel("Date")
        ax.set_ylabel("Sales ($)")
        ax.legend()
        ax.grid(True)
        
        self.display_chart(fig, "trends")
        
        # Update text
        growth_rate = (monthly_sales['Sales'].iloc[-1] / monthly_sales['Sales'].iloc[0] - 1) * 100
        volatility = monthly_sales['Sales'].std() / monthly_sales['Sales'].mean() * 100
        
        analysis = (
            "Market Trend Analysis\n\n"
            f"{'Observation Period:':<30}{monthly_sales.index[0].strftime('%b %Y')} to {monthly_sales.index[-1].strftime('%b %Y')}\n"
            f"{'Total Growth:':<30}{growth_rate:.2f}%\n"
            f"{'Average Monthly Sales:':<30}${monthly_sales['Sales'].mean():,.2f}\n"
            f"{'Market Volatility:':<30}{volatility:.2f}%\n"
            f"{'Current Trend (3-month MA):':<30}{'Upward' if monthly_sales['MA_3'].iloc[-1] > monthly_sales['MA_3'].iloc[-2] else 'Downward'}\n"
        )
        
        self.update_trends_text(analysis)
    
    def analyze_seasonality(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
            
        if 'Date' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Date' and 'Sales' not found.")
            return
            
        # Prepare data
        df_season = self.df[['Date', 'Sales']].copy()
        df_season['Month'] = df_season['Date'].dt.month
        df_season['Year'] = df_season['Date'].dt.year
        
        # Calculate seasonal indices
        monthly_avg = df_season.groupby('Month')['Sales'].mean()
        overall_avg = df_season['Sales'].mean()
        seasonal_index = (monthly_avg / overall_avg) * 100
        
        # Plot
        fig, ax = plt.subplots(figsize=(10, 6))
        months = [calendar.month_abbr[i] for i in range(1, 13)]
        ax.bar(months, seasonal_index, color='#2196F3')
        ax.axhline(100, color='red', linestyle='--', label='Average (100)')
        ax.set_title("Seasonal Sales Patterns")
        ax.set_xlabel("Month")
        ax.set_ylabel("Seasonal Index (%)")
        ax.legend()
        ax.grid(True)
        
        self.display_chart(fig, "trends")
        
        # Update text
        peak_month = months[np.argmax(seasonal_index) - 1]
        low_month = months[np.argmin(seasonal_index) - 1]
        peak_value = np.max(seasonal_index)
        low_value = np.min(seasonal_index)
        
        analysis = (
            "Seasonality Analysis\n\n"
            f"{'Peak Season:':<25}{peak_month} ({peak_value:.1f}% of average)\n"
            f"{'Low Season:':<25}{low_month} ({low_value:.1f}% of average)\n"
            f"{'Seasonal Variation:':<25}{peak_value - low_value:.1f}%\n"
            f"{'Recommendations:':<25}Stock up in {peak_month}, promotions in {low_month}"
        )
        
        self.update_trends_text(analysis)
    
    def forecast_demand(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
            
        if 'Date' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Date' and 'Sales' not found.")
            return
            
        # Prepare data
        df_forecast = self.df[['Date', 'Sales']].copy()
        df_forecast.set_index('Date', inplace=True)
        monthly_sales = df_forecast.resample('M').sum()
        
        # Simple forecasting using moving average
        forecast_periods = 6
        forecast = monthly_sales['Sales'].rolling(window=3).mean().iloc[-1]
        
        # Generate future dates
        last_date = monthly_sales.index[-1]
        future_dates = [last_date + relativedelta(months=i) for i in range(1, forecast_periods + 1)]
        
        # Plot
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(monthly_sales.index, monthly_sales['Sales'], label='Historical Sales', marker='o')
        ax.plot(future_dates, [forecast] * forecast_periods, 
               label='Forecast', linestyle='--', color='red')
        ax.set_title(f"Demand Forecast (Next {forecast_periods} Months)")
        ax.set_xlabel("Date")
        ax.set_ylabel("Sales ($)")
        ax.legend()
        ax.grid(True)
        
        self.display_chart(fig, "trends")
        
        # Update text
        analysis = (
            "Demand Forecasting\n\n"
            f"{'Forecast Period:':<25}Next {forecast_periods} months\n"
            f"{'Expected Monthly Sales:':<25}${forecast:,.2f}\n"
            f"{'Total Expected Sales:':<25}${forecast * forecast_periods:,.2f}\n"
            f"{'Forecast Method:':<25}3-Month Moving Average\n"
            f"{'Recommendations:':<25}Adjust inventory and staffing accordingly"
        )
        
        self.update_trends_text(analysis)
    
    def competitive_analysis(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
            
        if 'Product' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Product' and 'Sales' not found.")
            return
            
        # Calculate market share
        product_sales = self.df.groupby('Product')['Sales'].sum()
        total_sales = product_sales.sum()
        market_share = (product_sales / total_sales * 100).sort_values(ascending=False)
        
        # Plot
        fig, ax = plt.subplots(figsize=(10, 6))
        market_share.head(10).plot(kind='bar', color='#FF9800', ax=ax)
        ax.set_title("Top 10 Products by Market Share")
        ax.set_xlabel("Product")
        ax.set_ylabel("Market Share (%)")
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "trends")
        
        # Update text
        top_product = market_share.index[0]
        top_share = market_share.iloc[0]
        top3_share = market_share.head(3).sum()
        
        analysis = (
            "Competitive Analysis\n\n"
            f"{'Total Products:':<25}{len(market_share)}\n"
            f"{'Market Leader:':<25}{top_product}\n"
            f"{'Market Share:':<25}{top_share:.1f}%\n"
            f"{'Top 3 Market Share:':<25}{top3_share:.1f}%\n"
            f"{'Recommendations:':<25}Focus marketing on {top_product}, diversify product range"
        )
        
        self.update_trends_text(analysis)
    
    def update_trends_text(self, text):
        self.trends_text.config(state='normal')
        self.trends_text.delete(1.0, tk.END)
        self.trends_text.insert(tk.END, text)
        self.trends_text.config(state='disabled')
        
        # ==================== HELPER METHODS ====================
    
    def display_chart(self, figure, tab_name):
        """Display a matplotlib chart in the specified tab's chart frame"""
        # Clear previous chart
        frame = getattr(self, f"{tab_name}_chart_frame")
        for widget in frame.winfo_children():
            widget.destroy()
            
        # Embed new chart
        canvas = FigureCanvasTkAgg(figure, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Add toolbar
        toolbar = NavigationToolbar2Tk(canvas, frame)
        toolbar.update()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    
    # ==================== DATA MANAGEMENT METHODS ====================
    
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        
        try:
            self.df = pd.read_excel(file_path)
            numeric_cols = self.df.select_dtypes(include=np.number).columns.tolist()
            
            if not numeric_cols:
                messagebox.showerror("Error", "No numeric columns found in the file.")
                return
            
            self.column_dropdown['values'] = self.df.columns.tolist()
            self.column_dropdown.current(0)
            self.export_btn.config(state=tk.NORMAL)
            
            # Auto-detect date columns
            date_cols = [col for col in self.df.columns if self.df[col].dtype == 'datetime64[ns]']
            if date_cols:
                self.df['Date'] = self.df[date_cols[0]]  # Use first date column as default
            else:
                # Try to parse dates if they're in string format
                for col in self.df.columns:
                    if self.df[col].dtype == 'object':
                        try:
                            self.df['Date'] = pd.to_datetime(self.df[col])
                            break
                        except:
                            continue
            
            self.status_var.set(f"Loaded: {file_path} | Rows: {len(self.df)} | Columns: {len(self.df.columns)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
    
    def export_data(self):
        if self.df is None or self.df.empty:
            messagebox.showerror("Error", "No data to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save data as"
        )
        
        if not file_path:
            return
        
        try:
            self.df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Data exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data:\n{str(e)}")
    
    def fill_missing_with_mean(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        try:
            self.df.fillna(self.df.mean(numeric_only=True), inplace=True)
            messagebox.showinfo("Success", "Missing values filled with column means.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill missing values:\n{str(e)}")
    
    def fill_missing_with_zero(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        try:
            self.df.fillna(0, inplace=True)
            messagebox.showinfo("Success", "Missing values filled with zeros.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill missing values:\n{str(e)}")
    
    def drop_missing_rows(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        try:
            before = len(self.df)
            self.df.dropna(inplace=True)
            after = len(self.df)
            messagebox.showinfo("Success", f"Dropped {before - after} rows with missing values.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to drop rows:\n{str(e)}")
    
    def drop_duplicates(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        try:
            before = len(self.df)
            self.df.drop_duplicates(inplace=True)
            after = len(self.df)
            messagebox.showinfo("Success", f"Dropped {before - after} duplicate rows.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to drop duplicates:\n{str(e)}")
    
    def reset_index(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        try:
            self.df.reset_index(drop=True, inplace=True)
            messagebox.showinfo("Success", "Index reset successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to reset index:\n{str(e)}")
    
    def show_missing_summary(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        missing_counts = self.df.isna().sum()
        missing_percent = (missing_counts / len(self.df)) * 100
        summary = pd.DataFrame({
            'Missing Count': missing_counts,
            'Missing Percentage': missing_percent
        })
        
        popup = tk.Toplevel(self.root)
        popup.title("Missing Data Summary")
        popup.geometry("500x400")
        
        text_area = tk.Text(popup, wrap="none", font=("Consolas", 13),borderwidth=0,highlightthickness=0,bg="#1e1e1e",fg="white")
        text_area.pack(expand=True, fill="both")
        
        text_area.insert(tk.END, summary.to_string())
        text_area.config(state="disabled")

        # adding scrollbar to text
        scrollbar = ttk.Scrollbar(text_area,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.sales_metrics_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_area.yview)
        
    
    def calculate_stats(self):
        selected_col = self.column_var.get()
        if not selected_col or self.df is None:
            messagebox.showerror("Error", "No column selected or data loaded.")
            return
        
        data = self.df[selected_col].dropna()
        
        if data.empty:
            messagebox.showerror("Error", "Selected column has no valid numeric data.")
            return
        
        count = data.count()
        mean = data.mean()
        median = data.median()
        mode = stats.mode(data, keepdims=False).mode
        min_val = data.min()
        max_val = data.max()
        std_dev = data.std()
        variance = data.var()
        skewness = data.skew()
        kurtosis = data.kurtosis()
        p25 = np.percentile(data, 25)
        p75 = np.percentile(data, 75)
        moment2 = stats.moment(data, moment=2)
        moment3 = stats.moment(data, moment=3)
        
        result = (
            f"Analyzing: {selected_col}\n\n"
            f"{'Count:':<25}{count:.0f}\n"
            f"{'Mean:':<25}{mean:.2f}\n"
            f"{'Median:':<25}{median:.2f}\n"
            f"{'Mode:':<25}{mode}\n"
            f"{'Min:':<25}{min_val:.2f}\n"
            f"{'Max:':<25}{max_val:.2f}\n"
            f"{'Std Deviation:':<25}{std_dev:.2f}\n"
            f"{'Variance:':<25}{variance:.2f}\n"
            f"{'Skewness:':<25}{skewness:.2f}\n"
            f"{'Kurtosis:':<25}{kurtosis:.2f}\n"
            f"{'25th Percentile:':<25}{p25:.2f}\n"
            f"{'75th Percentile:':<25}{p75:.2f}\n"
            f"{'2nd Central Moment:':<25}{moment2:.2f}\n"
            f"{'3rd Central Moment:':<25}{moment3:.2f}"
        )
        
        self.result_box.config(state='normal')
        self.result_box.delete(1.0, tk.END)
        self.result_box.insert(tk.END, result)
        self.result_box.config(state='disabled')
    
    def plot_histogram(self):
        if self.df is None: return
        numeric_cols = self.df.select_dtypes(include='number').columns.tolist()
        if not numeric_cols:
            messagebox.showinfo("Info", "No numeric columns found.")
            return

        def plot():
            col = combo.get()
            plt.figure(figsize=(6, 4))
            sns.histplot(self.df[col], kde=True, color="skyblue")
            plt.title(f"Histogram of {col}")
            plt.tight_layout()
            plt.show()
            top.destroy()

        top = tk.Toplevel(self.root)
        top.title("Select Column for Histogram")
        combo = ttk.Combobox(top, values=numeric_cols)
        combo.pack(pady=10)
        tk.Button(top, text="Plot", command=plot).pack()

    def plot_boxplot(self):
        if self.df is None: return
        numeric_cols = self.df.select_dtypes(include='number').columns.tolist()
        if not numeric_cols:
            messagebox.showinfo("Info", "No numeric columns found.")
            return

        def plot():
            col = combo.get()
            plt.figure(figsize=(6, 4))
            sns.boxplot(x=self.df[col], color="orange")
            plt.title(f"Boxplot of {col}")
            plt.tight_layout()
            plt.show()
            top.destroy()

        top = tk.Toplevel(self.root)
        top.title("Select Column for Boxplot")
        combo = ttk.Combobox(top, values=numeric_cols)
        combo.pack(pady=10)
        tk.Button(top, text="Plot", command=plot).pack()
    
    def show_filter_dialog(self):
        selected_col = self.column_var.get()
        if not selected_col or self.df is None:
            messagebox.showerror("Error", "No column selected or data loaded.")
            return
        
        threshold = simpledialog.askfloat("Filter Data", f"Show rows where {selected_col} >", parent=self.root)
        if threshold is None:
            return
        
        filtered_df = self.df[self.df[selected_col] > threshold]
        
        if filtered_df.empty:
            messagebox.showinfo("No Data", f"No rows found where {selected_col} > {threshold}.")
            return
        
        popup = tk.Toplevel(self.root)
        popup.title(f"Filtered Data: {selected_col} > {threshold}")
        popup.geometry("800x500")
        
        text_area = tk.Text(popup, wrap="none", font=("Consolas", 13),borderwidth=0,highlightthickness=0,bg="#1e1e1e",fg="white")
        text_area.pack(expand=True, fill="both")
        
        text_area.insert(tk.END, filtered_df.to_string())
        text_area.config(state="disabled")

        # adding scrollbar to the text
        scrollbar = ttk.Scrollbar(text_area,orient="vertical",style="Custom.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.result_box.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_area.yview)
    
    # ==================== SALES ANALYSIS METHODS ====================
    
    def show_sales_trends(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Date' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Date' and 'Sales' not found.")
            return
        
        period = self.period_var.get()
        
        # Group by time period
        df_sales = self.df[['Date', 'Sales']].copy()
        df_sales.set_index('Date', inplace=True)
        
        if period == "Daily":
            df_grouped = df_sales.resample('D').sum()
            title = "Daily Sales Trends"
        elif period == "Weekly":
            df_grouped = df_sales.resample('W').sum()
            title = "Weekly Sales Trends"
        elif period == "Monthly":
            df_grouped = df_sales.resample('M').sum()
            title = "Monthly Sales Trends"
        elif period == "Quarterly":
            df_grouped = df_sales.resample('Q').sum()
            title = "Quarterly Sales Trends"
        else:  # Yearly
            df_grouped = df_sales.resample('Y').sum()
            title = "Yearly Sales Trends"
        
        # Plot
        fig, ax = plt.subplots(figsize=(5, 2))
        ax.plot(df_grouped.index, df_grouped['Sales'], marker='o', color='#4CAF50')
        ax.set_title(title)
        ax.set_xlabel("Date")
        ax.set_ylabel("Sales")
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "sales")
        
        # Update metrics
        metrics = (
            f"Sales Trends Analysis ({period})\n\n"
            f"{'Total Sales:':<20}${df_grouped['Sales'].sum():,.2f}\n"
            f"{'Average Sales:':<20}${df_grouped['Sales'].mean():,.2f}\n"
            f"{'Highest Sales:':<20}${df_grouped['Sales'].max():,.2f}\n"
            f"{'Lowest Sales:':<20}${df_grouped['Sales'].min():,.2f}\n"
            f"{'Growth Rate:':<20}{(df_grouped['Sales'].pct_change().mean() * 100):.2f}%"
        )
        
        self.sales_metrics_text.config(state='normal')
        self.sales_metrics_text.delete(1.0, tk.END)
        self.sales_metrics_text.insert(tk.END, metrics)
        self.sales_metrics_text.config(state='disabled')
    
    def show_product_performance(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Product' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Product' and 'Sales' not found.")
            return
        
        # Group by product
        df_product = self.df.groupby('Product')['Sales'].agg(['sum', 'count', 'mean']).sort_values('sum', ascending=False)
        df_product.columns = ['Total Sales', 'Number of Transactions', 'Average Sale']
        
        # Plot
        fig, ax = plt.subplots(figsize=(5, 2))
        df_product['Total Sales'].head(10).plot(kind='bar', color='#2196F3', ax=ax)
        ax.set_title("Top 10 Products by Sales")
        ax.set_xlabel("Product")
        ax.set_ylabel("Total Sales ($)")
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "sales")
        
        # Update metrics
        metrics = "Product Performance Analysis\n\n"
        metrics += f"{'Total Products:':<25}{len(df_product)}\n"
        metrics += f"{'Top Product:':<25}{df_product.index[0]}\n"
        metrics += f"{'Top Product Sales:':<25}${df_product.iloc[0]['Total Sales']:,.2f}\n"
        metrics += f"{'Avg Sale per Product:':<25}${df_product['Average Sale'].mean():,.2f}\n"
        metrics += f"{'Bottom Product:':<25}{df_product.index[-1]}\n"
        metrics += f"{'Bottom Product Sales:':<25}${df_product.iloc[-1]['Total Sales']:,.2f}"
        
        self.sales_metrics_text.config(state='normal')
        self.sales_metrics_text.delete(1.0, tk.END)
        self.sales_metrics_text.insert(tk.END, metrics)
        self.sales_metrics_text.config(state='disabled')
    
    def show_regional_sales(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Region' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Region' and 'Sales' not found.")
            return
        
        # Group by region
        df_region = self.df.groupby('Region')['Sales'].agg(['sum', 'count', 'mean']).sort_values('sum', ascending=False)
        df_region.columns = ['Total Sales', 'Number of Transactions', 'Average Sale']
        
        # Plot
        fig, ax = plt.subplots(figsize=(5, 2))
        df_region['Total Sales'].plot(kind='bar', color='#9C27B0', ax=ax)
        ax.set_title("Regional Sales Performance")
        ax.set_xlabel("Region")
        ax.set_ylabel("Total Sales ($)")
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "sales")
        
        # Update metrics
        metrics = "Regional Sales Analysis\n\n"
        metrics += f"{'Total Regions:':<25}{len(df_region)}\n"
        metrics += f"{'Top Region:':<25}{df_region.index[0]}\n"
        metrics += f"{'Top Region Sales:':<25}${df_region.iloc[0]['Total Sales']:,.2f}\n"
        metrics += f"{'Avg Sale per Region:':<25}${df_region['Average Sale'].mean():,.2f}\n"
        metrics += f"{'Bottom Region:':<25}{df_region.index[-1]}\n"
        metrics += f"{'Bottom Region Sales:':<25}${df_region.iloc[-1]['Total Sales']:,.2f}"
        
        self.sales_metrics_text.config(state='normal')
        self.sales_metrics_text.delete(1.0, tk.END)
        self.sales_metrics_text.insert(tk.END, metrics)
        self.sales_metrics_text.config(state='disabled')
    
    def show_revenue_distribution(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required column 'Sales' not found.")
            return
        
        # Plot distribution
        fig, ax = plt.subplots(figsize=(5, 2))
        sns.histplot(self.df['Sales'].dropna(), kde=True, color='#FF9800', ax=ax)
        ax.set_title("Revenue Distribution")
        ax.set_xlabel("Sales Amount ($)")
        ax.set_ylabel("Frequency")
        ax.grid(True)
        plt.tight_layout()
        
        self.display_chart(fig, "sales")
        
        # Calculate metrics
        sales = self.df['Sales'].dropna()
        skewness = sales.skew()
        kurt = sales.kurtosis()
        
        metrics = "Revenue Distribution Analysis\n\n"
        metrics += f"{'Total Sales:':<25}${sales.sum():,.2f}\n"
        metrics += f"{'Mean Sale:':<25}${sales.mean():,.2f}\n"
        metrics += f"{'Median Sale:':<25}${sales.median():,.2f}\n"
        metrics += f"{'Skewness:':<25}{skewness:.2f}\n"
        metrics += f"{'Kurtosis:':<25}{kurt:.2f}\n"
        metrics += f"{'95th Percentile:':<25}${np.percentile(sales, 95):,.2f}"
        
        self.sales_metrics_text.config(state='normal')
        self.sales_metrics_text.delete(1.0, tk.END)
        self.sales_metrics_text.insert(tk.END, metrics)
        self.sales_metrics_text.config(state='disabled')
    
    def show_peak_hours(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Date' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Date' and 'Sales' not found.")
            return
        
        # Extract hour from datetime
        df_sales = self.df.copy()
        df_sales['Hour'] = df_sales['Date'].dt.hour
        
        # Group by hour
        df_hourly = df_sales.groupby('Hour')['Sales'].agg(['sum', 'count']).sort_index()
        df_hourly.columns = ['Total Sales', 'Transaction Count']
        
        # Plot
        fig, ax = plt.subplots(figsize=(4, 2))
        df_hourly['Total Sales'].plot(kind='bar', color='#607D8B', ax=ax)
        ax.set_title("Sales by Hour of Day")
        ax.set_xlabel("Hour of Day")
        ax.set_ylabel("Total Sales ($)")
        ax.grid(True)
        plt.tight_layout()
        
        self.display_chart(fig, "sales")
        
        # Find peak hours
        peak_sales_hour = df_hourly['Total Sales'].idxmax()
        peak_trans_hour = df_hourly['Transaction Count'].idxmax()
        
        metrics = "Peak Hours Analysis\n\n"
        metrics += f"{'Peak Sales Hour:':<25}{peak_sales_hour}:00\n"
        metrics += f"{'Peak Sales Amount:':<25}${df_hourly.loc[peak_sales_hour, 'Total Sales']:,.2f}\n"
        metrics += f"{'Peak Transaction Hour:':<25}{peak_trans_hour}:00\n"
        metrics += f"{'Transactions at Peak:':<25}{df_hourly.loc[peak_trans_hour, 'Transaction Count']}"
        
        self.sales_metrics_text.config(state='normal')
        self.sales_metrics_text.delete(1.0, tk.END)
        self.sales_metrics_text.insert(tk.END, metrics)
        self.sales_metrics_text.config(state='disabled')
    
    def show_returns_analysis(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Returns' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Returns' and 'Sales' not found.")
            return
        
        # Calculate return rate
        total_sales = self.df['Sales'].sum()
        total_returns = self.df['Returns'].sum()
        return_rate = (total_returns / total_sales) * 100
        
        # Plot returns over time if date column exists
        fig, ax = plt.subplots(figsize=(5, 2))
        
        if 'Date' in self.df.columns:
            df_returns = self.df[['Date', 'Returns']].copy()
            df_returns.set_index('Date', inplace=True)
            df_returns = df_returns.resample('M').sum()
            
            ax.plot(df_returns.index, df_returns['Returns'], color='#E91E63', marker='o')
            ax.set_title("Monthly Returns")
            ax.set_xlabel("Date")
            ax.set_ylabel("Returns ($)")
        else:
            ax.bar(['Returns'], [total_returns], color='#E91E63')
            ax.set_title("Total Returns")
            ax.set_ylabel("Amount ($)")
        
        ax.grid(True)
        plt.tight_layout()
        
        self.display_chart(fig, "sales")
        
        # Update metrics
        metrics = "Returns Analysis\n\n"
        metrics += f"{'Total Sales:':<25}${total_sales:,.2f}\n"
        metrics += f"{'Total Returns:':<25}${total_returns:,.2f}\n"
        metrics += f"{'Return Rate:':<25}{return_rate:.2f}%\n"
        metrics += f"{'Net Revenue:':<25}${total_sales - total_returns:,.2f}"
        
        self.sales_metrics_text.config(state='normal')
        self.sales_metrics_text.delete(1.0, tk.END)
        self.sales_metrics_text.insert(tk.END, metrics)
        self.sales_metrics_text.config(state='disabled')
    
    # ==================== INVENTORY ANALYSIS METHODS ====================
    
    def calculate_reorder_points(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Product' not in self.df.columns or 'Inventory' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Product', 'Inventory', and 'Sales' not found.")
            return
        
        # Calculate reorder points (simplified example)
        df_inventory = self.df.groupby('Product').agg({
            'Inventory': 'mean',
            'Sales': 'mean'
        })
        
        # Assume lead time of 7 days and safety stock of 2 days
        df_inventory['Reorder Point'] = (df_inventory['Sales'] / 30 * 9).round()
        df_inventory['Status'] = np.where(
            df_inventory['Inventory'] <= df_inventory['Reorder Point'],
            'Reorder Now',
            'Stock OK'
        )
        
        # Plot
        fig, ax = plt.subplots(figsize=(8, 5))
        df_inventory['Inventory'].plot(kind='bar', color='#4CAF50', ax=ax, label='Current Inventory')
        df_inventory['Reorder Point'].plot(kind='line', color='red', marker='o', ax=ax, label='Reorder Point')
        ax.set_title("Inventory vs Reorder Points")
        ax.set_xlabel("Product")
        ax.set_ylabel("Quantity")
        ax.legend()
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "inventory")
        
        # Update metrics
        need_reorder = (df_inventory['Status'] == 'Reorder Now').sum()
        
        metrics = "Inventory Reorder Analysis\n\n"
        metrics += f"{'Total Products:':<25}{len(df_inventory)}\n"
        metrics += f"{'Products Need Reorder:':<25}{need_reorder}\n"
        metrics += f"{'Average Inventory:':<25}{df_inventory['Inventory'].mean():.1f}\n"
        metrics += f"{'Average Daily Sales:':<25}{df_inventory['Sales'].mean()/30:.1f}\n"
        metrics += f"{'Highest Reorder Point:':<25}{df_inventory['Reorder Point'].max():.0f}"
        
        self.inventory_metrics_text.config(state='normal')
        self.inventory_metrics_text.delete(1.0, tk.END)
        self.inventory_metrics_text.insert(tk.END, metrics)
        self.inventory_metrics_text.config(state='disabled')
    
    def show_stock_aging(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Product' not in self.df.columns or 'Date' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Product' and 'Date' not found.")
            return
        
        # Calculate days in inventory
        current_date = datetime.now()
        df_aging = self.df.copy()
        df_aging['Days in Stock'] = (current_date - df_aging['Date']).dt.days
        
        # Group by product
        df_aging = df_aging.groupby('Product')['Days in Stock'].mean().sort_values(ascending=False)
        
        # Plot
        fig, ax = plt.subplots(figsize=(5,3))
        df_aging.head(10).plot(kind='bar', color='#2196F3', ax=ax)
        ax.set_title("Top 10 Products by Days in Stock")
        ax.set_xlabel("Product")
        ax.set_ylabel("Average Days in Stock")
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "inventory")
        
        # Update metrics
        metrics = "Stock Aging Analysis\n\n"
        metrics += f"{'Oldest Product:':<25}{df_aging.index[0]}\n"
        metrics += f"{'Days in Stock:':<25}{df_aging.iloc[0]:.1f}\n"
        metrics += f"{'Average Days:':<25}{df_aging.mean():.1f}\n"
        metrics += f"{'Newest Product:':<25}{df_aging.index[-1]}\n"
        metrics += f"{'Days in Stock:':<25}{df_aging.iloc[-1]:.1f}"
        
        self.inventory_metrics_text.config(state='normal')
        self.inventory_metrics_text.delete(1.0, tk.END)
        self.inventory_metrics_text.insert(tk.END, metrics)
        self.inventory_metrics_text.config(state='disabled')
    
    def show_stock_levels(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Product' not in self.df.columns or 'Inventory' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Product' and 'Inventory' not found.")
            return
        
        # Calculate over/under stock (simplified example)
        df_stock = self.df.groupby('Product')['Inventory'].mean().sort_values(ascending=False)
        median_stock = df_stock.median()
        
        df_stock = pd.DataFrame(df_stock)
        df_stock['Status'] = np.where(
            df_stock['Inventory'] > median_stock * 1.5,
            'Overstocked',
            np.where(
                df_stock['Inventory'] < median_stock * 0.5,
                'Understocked',
                'Normal'
            )
        )
        
        # Plot
        fig, ax = plt.subplots(figsize=(8, 5))
        colors = {'Overstocked': '#E91E63', 'Understocked': '#FF9800', 'Normal': '#4CAF50'}
        
        for status, color in colors.items():
            subset = df_stock[df_stock['Status'] == status]
            ax.bar(subset.index, subset['Inventory'], color=color, label=status)
        
        ax.axhline(median_stock * 1.5, color='red', linestyle='--', label='Overstock Threshold')
        ax.axhline(median_stock * 0.5, color='orange', linestyle='--', label='Understock Threshold')
        ax.set_title("Stock Levels Analysis")
        ax.set_xlabel("Product")
        ax.set_ylabel("Inventory Level")
        ax.legend()
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "inventory")
        
        # Update metrics
        overstocked = (df_stock['Status'] == 'Overstocked').sum()
        understocked = (df_stock['Status'] == 'Understocked').sum()
        
        metrics = "Stock Levels Analysis\n\n"
        metrics += f"{'Total Products:':<25}{len(df_stock)}\n"
        metrics += f"{'Overstocked:':<25}{overstocked}\n"
        metrics += f"{'Understocked:':<25}{understocked}\n"
        metrics += f"{'Median Inventory:':<25}{median_stock:.1f}\n"
        metrics += f"{'Max Inventory:':<25}{df_stock['Inventory'].max():.1f}"
        
        self.inventory_metrics_text.config(state='normal')
        self.inventory_metrics_text.delete(1.0, tk.END)
        self.inventory_metrics_text.insert(tk.END, metrics)
        self.inventory_metrics_text.config(state='disabled')
    
    def calculate_inventory_turnover(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Product' not in self.df.columns or 'Inventory' not in self.df.columns or 'Sales' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Product', 'Inventory', and 'Sales' not found.")
            return
        
        # Calculate inventory turnover (simplified example)
        df_turnover = self.df.groupby('Product').agg({
            'Inventory': 'mean',
            'Sales': 'sum'
        })
        
        df_turnover['Turnover'] = df_turnover['Sales'] / df_turnover['Inventory']
        
        # Plot
        fig, ax = plt.subplots(figsize=(8, 5))
        df_turnover['Turnover'].sort_values(ascending=False).head(10).plot(
            kind='bar', color='#9C27B0', ax=ax)
        ax.set_title("Top 10 Products by Inventory Turnover")
        ax.set_xlabel("Product")
        ax.set_ylabel("Turnover Ratio")
        ax.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        self.display_chart(fig, "inventory")
        
        # Update metrics
        avg_turnover = df_turnover['Turnover'].mean()
        fastest = df_turnover['Turnover'].idxmax()
        slowest = df_turnover['Turnover'].idxmin()
        
        metrics = "Inventory Turnover Analysis\n\n"
        metrics += f"{'Average Turnover:':<25}{avg_turnover:.2f}\n"
        metrics += f"{'Fastest Turning:':<25}{fastest}\n"
        metrics += f"{'Turnover Rate:':<25}{df_turnover.loc[fastest, 'Turnover']:.2f}\n"
        metrics += f"{'Slowest Turning:':<25}{slowest}\n"
        metrics += f"{'Turnover Rate:':<25}{df_turnover.loc[slowest, 'Turnover']:.2f}"
        
        self.inventory_metrics_text.config(state='normal')
        self.inventory_metrics_text.delete(1.0, tk.END)
        self.inventory_metrics_text.insert(tk.END, metrics)
        self.inventory_metrics_text.config(state='disabled')
    
    # ==================== FINANCIAL ANALYSIS METHODS ====================
    
    def show_profit_analysis(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Revenue' not in self.df.columns or 'Cost' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Revenue' and 'Cost' not found.")
            return
        
        # Calculate profit metrics
        self.df['Profit'] = self.df['Revenue'] - self.df['Cost']
        self.df['Profit Margin'] = (self.df['Profit'] / self.df['Revenue']) * 100
        
        # Group by time period
        period = self.finance_period_var.get()
        df_finance = self.df[['Date', 'Revenue', 'Cost', 'Profit', 'Profit Margin']].copy()
        df_finance.set_index('Date', inplace=True)
        
        if period == "Month-over-Month":
            df_grouped = df_finance.resample('M').sum()
            title = "Monthly Profit Analysis"
        elif period == "Quarter-over-Quarter":
            df_grouped = df_finance.resample('Q').sum()
            title = "Quarterly Profit Analysis"
        else:  # Year-over-Year
            df_grouped = df_finance.resample('Y').sum()
            title = "Yearly Profit Analysis"
        
        # Plot
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(5,5))
        
        # Revenue, Cost, Profit
        df_grouped[['Revenue', 'Cost', 'Profit']].plot(kind='bar', ax=ax1)
        ax1.set_title(f"{title} - Amounts")
        ax1.set_ylabel("Amount ($)")
        ax1.grid(True)
        
        # Profit Margin
        df_grouped['Profit Margin'].plot(kind='line', marker='o', color='green', ax=ax2)
        ax2.set_title("Profit Margin")
        ax2.set_ylabel("Percentage (%)")
        ax2.grid(True)
        
        plt.tight_layout()
        
        self.display_chart(fig, "finance")
        
        # Update metrics
        total_profit = df_grouped['Profit'].sum()
        avg_margin = df_grouped['Profit Margin'].mean()
        best_period = df_grouped['Profit'].idxmax().strftime('%Y-%m')
        best_profit = df_grouped['Profit'].max()
        
        metrics = f"Profit Analysis ({period})\n\n"
        metrics += f"{'Total Profit:':<25}${total_profit:,.2f}\n"
        metrics += f"{'Average Margin:':<25}{avg_margin:.2f}%\n"
        metrics += f"{'Best Period:':<25}{best_period}\n"
        metrics += f"{'Best Profit:':<25}${best_profit:,.2f}\n"
        metrics += f"{'Revenue Growth:':<25}{df_grouped['Revenue'].pct_change().mean()*100:.2f}%\n"
        metrics += f"{'Profit Growth:':<25}{df_grouped['Profit'].pct_change().mean()*100:.2f}%"
        
        self.finance_metrics_text.config(state='normal')
        self.finance_metrics_text.delete(1.0, tk.END)
        self.finance_metrics_text.insert(tk.END, metrics)
        self.finance_metrics_text.config(state='disabled')
    
    def show_growth_tracking(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Revenue' not in self.df.columns or 'Date' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Revenue' and 'Date' not found.")
            return
        
        # Calculate growth metrics
        period = self.finance_period_var.get()
        df_growth = self.df[['Date', 'Revenue']].copy()
        df_growth.set_index('Date', inplace=True)
        
        if period == "Month-over-Month":
            df_grouped = df_growth.resample('M').sum()
            title = "Monthly Growth Tracking"
        elif period == "Quarter-over-Quarter":
            df_grouped = df_growth.resample('Q').sum()
            title = "Quarterly Growth Tracking"
        else:  # Year-over-Year
            df_grouped = df_growth.resample('Y').sum()
            title = "Yearly Growth Tracking"
        
        df_grouped['Growth'] = df_grouped['Revenue'].pct_change() * 100
        df_grouped['Cumulative Growth'] = (df_grouped['Revenue'] / df_grouped['Revenue'].iloc[0] - 1) * 100
        
        # Plot
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 8))
        
        # Revenue
        df_grouped['Revenue'].plot(kind='bar', color='#2196F3', ax=ax1)
        ax1.set_title(f"{title} - Revenue")
        ax1.set_ylabel("Revenue ($)")
        ax1.grid(True)
        
        # Growth
        df_grouped['Growth'].plot(kind='line', marker='o', color='#4CAF50', ax=ax2)
        ax2.axhline(0, color='red', linestyle='--')
        ax2.set_title("Growth Rate")
        ax2.set_ylabel("Percentage Change (%)")
        ax2.grid(True)
        
        plt.tight_layout()
        
        self.display_chart(fig, "finance")
        
        # Update metrics
        total_growth = df_grouped['Cumulative Growth'].iloc[-1]
        avg_growth = df_grouped['Growth'].mean()
        best_growth = df_grouped['Growth'].max()
        worst_growth = df_grouped['Growth'].min()
        
        metrics = f"Growth Tracking ({period})\n\n"
        metrics += f"{'Total Growth:':<25}{total_growth:.2f}%\n"
        metrics += f"{'Average Growth:':<25}{avg_growth:.2f}%\n"
        metrics += f"{'Highest Growth:':<25}{best_growth:.2f}%\n"
        metrics += f"{'Lowest Growth:':<25}{worst_growth:.2f}%\n"
        metrics += f"{'Current Revenue:':<25}${df_grouped['Revenue'].iloc[-1]:,.2f}"
        
        self.finance_metrics_text.config(state='normal')
        self.finance_metrics_text.delete(1.0, tk.END)
        self.finance_metrics_text.insert(tk.END, metrics)
        self.finance_metrics_text.config(state='disabled')
    
    def show_financial_ratios(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Revenue' not in self.df.columns or 'Cost' not in self.df.columns or 'Inventory' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Revenue', 'Cost', and 'Inventory' not found.")
            return
        
        # Calculate financial ratios
        df_ratios = self.df[['Date', 'Revenue', 'Cost', 'Inventory']].copy()
        df_ratios['Profit'] = df_ratios['Revenue'] - df_ratios['Cost']
        
        period = self.finance_period_var.get()
        if period == "Month-over-Month":
            df_grouped = df_ratios.resample('M', on='Date').sum()
            title = "Monthly Financial Ratios"
        elif period == "Quarter-over-Quarter":
            df_grouped = df_ratios.resample('Q', on='Date').sum()
            title = "Quarterly Financial Ratios"
        else:  # Year-over-Year
            df_grouped = df_ratios.resample('Y', on='Date').sum()
            title = "Yearly Financial Ratios"
        
        df_grouped['Gross Margin'] = (df_grouped['Profit'] / df_grouped['Revenue']) * 100
        df_grouped['Operating Margin'] = (df_grouped['Profit'] / df_grouped['Revenue']) * 100  # Simplified
        df_grouped['Inventory Turnover'] = df_grouped['Cost'] / df_grouped['Inventory']
        
        # Plot
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 8))
        
        # Margins
        df_grouped[['Gross Margin', 'Operating Margin']].plot(kind='line', ax=ax1)
        ax1.set_title(f"{title} - Margins")
        ax1.set_ylabel("Percentage (%)")
        ax1.grid(True)
        
        # Inventory Turnover
        df_grouped['Inventory Turnover'].plot(kind='bar', color='#9C27B0', ax=ax2)
        ax2.set_title("Inventory Turnover")
        ax2.set_ylabel("Turnover Ratio")
        ax2.grid(True)
        
        plt.tight_layout()
        
        self.display_chart(fig, "finance")
        
        # Update metrics
        avg_gross = df_grouped['Gross Margin'].mean()
        avg_operating = df_grouped['Operating Margin'].mean()
        avg_turnover = df_grouped['Inventory Turnover'].mean()
        
        metrics = f"Financial Ratios ({period})\n\n"
        metrics += f"{'Avg Gross Margin:':<25}{avg_gross:.2f}%\n"
        metrics += f"{'Avg Operating Margin:':<25}{avg_operating:.2f}%\n"
        metrics += f"{'Avg Inventory Turnover:':<25}{avg_turnover:.2f}\n"
        metrics += f"{'Current Gross Margin:':<25}{df_grouped['Gross Margin'].iloc[-1]:.2f}%\n"
        metrics += f"{'Current Turnover:':<25}{df_grouped['Inventory Turnover'].iloc[-1]:.2f}"
        
        self.finance_metrics_text.config(state='normal')
        self.finance_metrics_text.delete(1.0, tk.END)
        self.finance_metrics_text.insert(tk.END, metrics)
        self.finance_metrics_text.config(state='disabled')
    
    def show_expense_analysis(self):
        if not hasattr(self, 'df') or self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        if 'Cost' not in self.df.columns or 'Date' not in self.df.columns:
            messagebox.showerror("Error", "Required columns 'Cost' and 'Date' not found.")
            return
        
        # Calculate expense metrics
        period = self.finance_period_var.get()
        df_expenses = self.df[['Date', 'Cost']].copy()
        df_expenses.set_index('Date', inplace=True)
        
        if period == "Month-over-Month":
            df_grouped = df_expenses.resample('M').sum()
            title = "Monthly Expense Analysis"
        elif period == "Quarter-over-Quarter":
            df_grouped = df_expenses.resample('Q').sum()
            title = "Quarterly Expense Analysis"
        else:  # Year-over-Year
            df_grouped = df_expenses.resample('Y').sum()
            title = "Yearly Expense Analysis"
        
        df_grouped['Growth'] = df_grouped['Cost'].pct_change() * 100
        
        # Plot
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 8))
        
        # Expenses
        df_grouped['Cost'].plot(kind='bar', color='#FF9800', ax=ax1)
        ax1.set_title(f"{title} - Expenses")
        ax1.set_ylabel("Amount ($)")
        ax1.grid(True)
        
        # Growth
        df_grouped['Growth'].plot(kind='line', marker='o', color='#E91E63', ax=ax2)
        ax2.axhline(0, color='red', linestyle='--')
        ax2.set_title("Expense Growth")
        ax2.set_ylabel("Percentage Change (%)")
        ax2.grid(True)
        
        plt.tight_layout()
        
        self.display_chart(fig, "finance")
        
        # Update metrics
        total_expenses = df_grouped['Cost'].sum()
        avg_growth = df_grouped['Growth'].mean()
        highest_expense = df_grouped['Cost'].max()
        lowest_expense = df_grouped['Cost'].min()
        
        metrics = f"Expense Analysis ({period})\n\n"
        metrics += f"{'Total Expenses:':<25}${total_expenses:,.2f}\n"
        metrics += f"{'Average Growth:':<25}{avg_growth:.2f}%\n"
        metrics += f"{'Highest Expense:':<25}${highest_expense:,.2f}\n"
        metrics += f"{'Lowest Expense:':<25}${lowest_expense:,.2f}"
        
        self.finance_metrics_text.config(state='normal')
        self.finance_metrics_text.delete(1.0, tk.END)
        self.finance_metrics_text.insert(tk.END, metrics)
        self.finance_metrics_text.config(state='disabled')
    
    # ==================== REPORTING METHODS ====================
    
    def generate_report(self):
        report_type = self.report_type_var.get()
        period = self.report_period_var.get()
        
        if self.df is None:
            messagebox.showerror("Error", "No data loaded.")
            return
        
        report = f"Business Report - {report_type}\n"
        report += f"Period: {period}\n"
        report += f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "="*50 + "\n\n"
        
        if report_type == "Sales Summary":
            report += self.generate_sales_report(period)
        elif report_type == "Inventory Status":
            report += self.generate_inventory_report()
        elif report_type == "Financial Summary":
            report += self.generate_financial_report(period)
        else:  # Comprehensive
            report += self.generate_sales_report(period) + "\n"
            report += self.generate_inventory_report() + "\n"
            report += self.generate_financial_report(period)
        
        self.report_text.config(state='normal')
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)
        self.report_text.config(state='disabled')
    
    def generate_sales_report(self, period):
        if 'Sales' not in self.df.columns or 'Date' not in self.df.columns:
            return "Sales data not available in the dataset.\n"
        
        df_sales = self.df[['Date', 'Sales']].copy()
        df_sales.set_index('Date', inplace=True)
        
        if period == "Daily":
            df_grouped = df_sales.resample('D').sum()
        elif period == "Weekly":
            df_grouped = df_sales.resample('W').sum()
        elif period == "Monthly":
            df_grouped = df_sales.resample('M').sum()
        elif period == "Quarterly":
            df_grouped = df_sales.resample('Q').sum()
        else:  # Yearly
            df_grouped = df_sales.resample('Y').sum()
        
        report = "SALES SUMMARY\n"
        report += "-"*50 + "\n"
        report += f"Total Sales: ${df_grouped['Sales'].sum():,.2f}\n"
        report += f"Average Sales: ${df_grouped['Sales'].mean():,.2f}\n"
        report += f"Highest Sales Period: {df_grouped['Sales'].idxmax().strftime('%Y-%m-%d')} (${df_grouped['Sales'].max():,.2f})\n"
        report += f"Lowest Sales Period: {df_grouped['Sales'].idxmin().strftime('%Y-%m-%d')} (${df_grouped['Sales'].min():,.2f})\n"
        report += f"Growth Rate: {(df_grouped['Sales'].pct_change().mean() * 100):.2f}%\n\n"
        
        # Add top products if available
        if 'Product' in self.df.columns:
            top_products = self.df.groupby('Product')['Sales'].sum().nlargest(5)
            report += "TOP 5 PRODUCTS BY SALES:\n"
            for product, sales in top_products.items():
                report += f"- {product}: ${sales:,.2f}\n"
        
        return report

    def generate_inventory_report(self):
        if 'Inventory' not in self.df.columns or 'Product' not in self.df.columns:
            return "Inventory data not available in the dataset.\n"
        
        report = "INVENTORY STATUS\n"
        report += "-"*50 + "\n"
        
        # Basic inventory metrics
        total_inventory = self.df['Inventory'].sum()
        avg_inventory = self.df['Inventory'].mean()
        report += f"Total Inventory Value: {total_inventory:,.0f} units\n"
        report += f"Average Inventory per Product: {avg_inventory:,.1f} units\n"
        
        # Identify low stock items
        low_stock_threshold = self.df['Inventory'].quantile(0.25)
        low_stock = self.df[self.df['Inventory'] <= low_stock_threshold]
        if not low_stock.empty:
            report += f"\nLOW STOCK ITEMS (≤ {low_stock_threshold:,.0f} units):\n"
            for _, row in low_stock.iterrows():
                report += f"- {row['Product']}: {row['Inventory']} units\n"
        
        # Identify overstock items
        overstock_threshold = self.df['Inventory'].quantile(0.75)
        overstock = self.df[self.df['Inventory'] >= overstock_threshold]
        if not overstock.empty:
            report += f"\nOVERSTOCK ITEMS (≥ {overstock_threshold:,.0f} units):\n"
            for _, row in overstock.iterrows():
                report += f"- {row['Product']}: {row['Inventory']} units\n"
        
        return report

    def generate_financial_report(self, period):
        if 'Revenue' not in self.df.columns or 'Cost' not in self.df.columns:
            return "Financial data not available in the dataset.\n"
        
        self.df['Profit'] = self.df['Revenue'] - self.df['Cost']
        self.df['Profit Margin'] = (self.df['Profit'] / self.df['Revenue']) * 100
        
        df_finance = self.df[['Date', 'Revenue', 'Cost', 'Profit', 'Profit Margin']].copy()
        df_finance.set_index('Date', inplace=True)
        
        if period == "Daily":
            df_grouped = df_finance.resample('D').sum()
        elif period == "Weekly":
            df_grouped = df_finance.resample('W').sum()
        elif period == "Monthly":
            df_grouped = df_finance.resample('M').sum()
        elif period == "Quarterly":
            df_grouped = df_finance.resample('Q').sum()
        else:  # Yearly
            df_grouped = df_finance.resample('Y').sum()
        
        report = "FINANCIAL SUMMARY\n"
        report += "-"*50 + "\n"
        report += f"Total Revenue: ${df_grouped['Revenue'].sum():,.2f}\n"
        report += f"Total Costs: ${df_grouped['Cost'].sum():,.2f}\n"
        report += f"Total Profit: ${df_grouped['Profit'].sum():,.2f}\n"
        report += f"Average Profit Margin: {df_grouped['Profit Margin'].mean():.2f}%\n"
        report += f"Highest Profit Period: {df_grouped['Profit'].idxmax().strftime('%Y-%m-%d')} (${df_grouped['Profit'].max():,.2f})\n"
        report += f"Lowest Profit Period: {df_grouped['Profit'].idxmin().strftime('%Y-%m-%d')} (${df_grouped['Profit'].min():,.2f})\n"
        
        return report

    def export_report(self):
        if self.report_text.get("1.0", "end-1c") == "":
            messagebox.showerror("Error", "No report generated to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save report as"
        )
        
        if not file_path:
            return
        
        try:
            # Create a DataFrame from the report text
            report_text = self.report_text.get("1.0", tk.END)
            df_report = pd.DataFrame([line for line in report_text.split('\n') if line.strip()])
            
            # Write to Excel
            with pd.ExcelWriter(file_path) as writer:
                df_report.to_excel(writer, sheet_name='Business Report', index=False, header=False)
            
            messagebox.showinfo("Success", f"Report exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report:\n{str(e)}")

    def export_pdf(self):
        messagebox.showinfo("Info", "PDF export functionality would be implemented here.")
        
    def analyze_market_trends(self):
        self.update_trends_text("Market trends analysis would be implemented here.")
        
    def analyze_seasonality(self):
        self.update_trends_text("Seasonality analysis would be implemented here.")
        
    def forecast_demand(self):
        self.update_trends_text("Demand forecasting would be implemented here.")
        
    def competitive_analysis(self):
        self.update_trends_text("Competitive analysis would be implemented here.")
        
    def update_trends_text(self, text):
        self.trends_text.config(state='normal')
        self.trends_text.delete(1.0, tk.END)
        self.trends_text.insert(tk.END, text)
        self.trends_text.config(state='disabled')
        
    def display_chart(self, figure, tab_name):
        # Clear previous chart
        frame = getattr(self, f"{tab_name}_chart_frame")
        for widget in frame.winfo_children():
            widget.destroy()
            
        # Embed new chart
        canvas = FigureCanvasTkAgg(figure, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Add toolbar
        toolbar = NavigationToolbar2Tk(canvas, frame)
        toolbar.update()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ==================== EXIT AND ABOUT US METHODS ====================
    
    def exit_application(self):
        """Exit the application with a confirmation dialog"""
        if messagebox.askyesno("Exit", "Are you sure you want to exit the application?"):
            self.root.destroy()
    
    def show_about_us(self):
        """Show information about the application and developers"""
        about_text = (
            "Business Management System\n"
            "Version: 1.0\n\n"
            "Developed by Group No. 102\n\n"
            "This comprehensive business management system provides:\n"
            "- Sales analysis and forecasting\n"
            "- Inventory management tools\n"
            "- Financial reporting and metrics\n"
            "- Market trend analysis\n"
            "- Custom report generation\n\n"
            "© 2025 Business Solutions. All rights reserved."
        )
        
        messagebox.showinfo("About Us", about_text)
 
 # Main application
if __name__ == "__main__":
    root = tk.Tk()
    app = BusinessManagementSystem(root)
    root.update_idletasks()
    root.mainloop()                     