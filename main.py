from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.popup import Popup
from kivy.metrics import dp
from kivy.core.window import Window

from openpyxl import Workbook, load_workbook
import os
from datetime import datetime, timedelta

# ---------------- SETTINGS ----------------

GST_RATE = 0.05
FILE_NAME = "GST_Billing_Data.xlsx"

# DARK THEME
BG_COLOR = (0.08, 0.08, 0.08, 1)
CARD_COLOR = (0.12, 0.12, 0.12, 1)
BTN_COLOR = (0.18, 0.55, 0.95, 1)
TEXT_COLOR = (1, 1, 1, 1)

# REMOVE DESKTOP WINDOW SIZE
Window.clearcolor = BG_COLOR


# ---------------- EXCEL SETUP ----------------

if not os.path.exists(FILE_NAME):
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Billing_Data"

    sheet.append([
        "Date",
        "Payment Mode",
        "Bill No",
        "Sale Amount",
        "Taxable Amount",
        "CGST",
        "SGST",
        "Total GST"
    ])

    wb.save(FILE_NAME)


# ---------------- GST CALCULATION ----------------

def calculate_gst(amount):
    taxable = amount / (1 + GST_RATE)
    gst = amount - taxable

    return (
        round(taxable, 2),
        round(gst / 2, 2),
        round(gst / 2, 2),
        round(gst, 2)
    )


# ---------------- MAIN APP ----------------

class GSTBillingApp(App):

    def build(self):

        self.current_date = ""
        self.phonepe_bill_no = 1
        self.cash_bill_no = 1

        self.total_sale = 0
        self.total_taxable = 0

        root = BoxLayout(
            orientation="vertical",
            padding=dp(12),
            spacing=dp(10)
        )

        # TITLE

        title = Label(
            text="GST BILLING SYSTEM",
            size_hint_y=None,
            height=dp(60),
            font_size="24sp",
            bold=True,
            color=TEXT_COLOR
        )

        root.add_widget(title)

        # SCROLL AREA

        scroll = ScrollView()

        self.layout = GridLayout(
            cols=1,
            spacing=dp(12),
            size_hint_y=None,
            padding=dp(5)
        )

        self.layout.bind(minimum_height=self.layout.setter('height'))

        scroll.add_widget(self.layout)

        root.add_widget(scroll)

        # ---------------- DATE INPUT ----------------

        self.date_input = TextInput(
            hint_text="Enter Date (DD-MM-YY)",
            multiline=False,
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=CARD_COLOR,
            foreground_color=TEXT_COLOR,
            cursor_color=TEXT_COLOR
        )

        self.layout.add_widget(self.date_input)

        # ---------------- BILL INPUTS ----------------

        self.phonepe_input = TextInput(
            hint_text="Starting PhonePe Bill Number",
            multiline=False,
            input_filter="int",
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=CARD_COLOR,
            foreground_color=TEXT_COLOR
        )

        self.layout.add_widget(self.phonepe_input)

        self.cash_input = TextInput(
            hint_text="Starting Cash Bill Number",
            multiline=False,
            input_filter="int",
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=CARD_COLOR,
            foreground_color=TEXT_COLOR
        )

        self.layout.add_widget(self.cash_input)

        # ---------------- PAYMENT MODE ----------------

        self.mode_spinner = Spinner(
            text="PhonePe",
            values=("PhonePe", "Cash"),
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=BTN_COLOR
        )

        self.layout.add_widget(self.mode_spinner)

        # ---------------- AMOUNT INPUT ----------------

        self.amount_input = TextInput(
            hint_text="Enter Sale Amount",
            multiline=False,
            input_filter="float",
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=CARD_COLOR,
            foreground_color=TEXT_COLOR
        )

        self.layout.add_widget(self.amount_input)

        # ---------------- ADD BILL BUTTON ----------------

        add_btn = Button(
            text="ADD BILL",
            size_hint_y=None,
            height=dp(58),
            font_size="20sp",
            bold=True,
            background_color=BTN_COLOR
        )

        add_btn.bind(on_press=self.add_bill)

        self.layout.add_widget(add_btn)

        # ---------------- EDIT BUTTON ----------------

        edit_btn = Button(
            text="EDIT BILL",
            size_hint_y=None,
            height=dp(58),
            font_size="20sp",
            bold=True,
            background_color=(0.85, 0.45, 0.1, 1)
        )

        edit_btn.bind(on_press=self.open_edit_popup)

        self.layout.add_widget(edit_btn)

        # ---------------- FINISH DAY BUTTON ----------------

        finish_btn = Button(
            text="FINISH DAY",
            size_hint_y=None,
            height=dp(58),
            font_size="20sp",
            bold=True,
            background_color=(0.2, 0.7, 0.3, 1)
        )

        finish_btn.bind(on_press=self.finish_day)

        self.layout.add_widget(finish_btn)

        # ---------------- SUMMARY LABEL ----------------

        self.summary_label = Label(
            text="Daily Totals Will Appear Here",
            size_hint_y=None,
            height=dp(160),
            font_size="18sp",
            halign="left",
            valign="top",
            color=TEXT_COLOR
        )

        self.summary_label.bind(size=self.summary_label.setter('text_size'))

        self.layout.add_widget(self.summary_label)

        return root

    # ---------------- SHOW POPUP ----------------

    def show_popup(self, title, message):

        popup_layout = BoxLayout(
            orientation="vertical",
            padding=dp(15),
            spacing=dp(15)
        )

        msg = Label(
            text=message,
            color=TEXT_COLOR
        )

        close_btn = Button(
            text="OK",
            size_hint_y=None,
            height=dp(50),
            background_color=BTN_COLOR
        )

        popup_layout.add_widget(msg)
        popup_layout.add_widget(close_btn)

        popup = Popup(
            title=title,
            content=popup_layout,
            size_hint=(0.85, 0.45),
            background_color=BG_COLOR
        )

        close_btn.bind(on_press=popup.dismiss)

        popup.open()

    # ---------------- ADD BILL ----------------

    def add_bill(self, instance):

        try:

            if self.current_date == "":

                self.current_date = self.date_input.text.strip()

                datetime.strptime(self.current_date, "%d-%m-%y")

                self.phonepe_bill_no = int(self.phonepe_input.text)
                self.cash_bill_no = int(self.cash_input.text)

            amount = float(self.amount_input.text)

            taxable, cgst, sgst, gst = calculate_gst(amount)

            mode = self.mode_spinner.text

            if mode == "PhonePe":
                bill_no = self.phonepe_bill_no
                self.phonepe_bill_no += 1
            else:
                bill_no = self.cash_bill_no
                self.cash_bill_no += 1

            self.total_sale += amount
            self.total_taxable += taxable

            wb = load_workbook(FILE_NAME)
            sheet = wb.active

            sheet.append([
                self.current_date,
                mode,
                bill_no,
                amount,
                taxable,
                cgst,
                sgst,
                gst
            ])

            wb.save(FILE_NAME)

            self.summary_label.text = (
                f"Last Bill Added\n\n"
                f"Mode: {mode}\n"
                f"Bill No: {bill_no}\n"
                f"Sale Amount: ₹ {amount:.2f}\n"
                f"Taxable: ₹ {taxable:.2f}\n"
                f"CGST: ₹ {cgst:.2f}\n"
                f"SGST: ₹ {sgst:.2f}\n"
                f"Total GST: ₹ {gst:.2f}\n\n"
                f"DAY SALE TOTAL: ₹ {self.total_sale:.2f}\n"
                f"DAY TAXABLE TOTAL: ₹ {self.total_taxable:.2f}"
            )

            self.amount_input.text = ""

        except Exception as e:
            self.show_popup("Error", str(e))

    # ---------------- FINISH DAY ----------------

    def finish_day(self, instance):

        try:

            wb = load_workbook(FILE_NAME)
            sheet = wb.active

            sheet.append(["", "", "", "", "", "", "", ""])

            wb.save(FILE_NAME)

            self.show_popup(
                "Day Finished",
                f"Date: {self.current_date}\n\n"
                f"Total Sale: ₹ {self.total_sale:.2f}\n"
                f"Total Taxable: ₹ {self.total_taxable:.2f}"
            )

            # AUTO INCREMENT DATE

            old_date = datetime.strptime(self.current_date, "%d-%m-%y")
            new_date = old_date + timedelta(days=1)

            self.current_date = new_date.strftime("%d-%m-%y")

            self.date_input.text = self.current_date

            self.total_sale = 0
            self.total_taxable = 0

        except Exception as e:
            self.show_popup("Error", str(e))

    # ---------------- EDIT BILL ----------------

    def open_edit_popup(self, instance):

        layout = BoxLayout(
            orientation="vertical",
            spacing=dp(10),
            padding=dp(10)
        )

        mode_spinner = Spinner(
            text="PhonePe",
            values=("PhonePe", "Cash"),
            size_hint_y=None,
            height=dp(50)
        )

        bill_input = TextInput(
            hint_text="Enter Bill Number",
            multiline=False,
            input_filter="int",
            size_hint_y=None,
            height=dp(50)
        )

        amount_input = TextInput(
            hint_text="New Sale Amount",
            multiline=False,
            input_filter="float",
            size_hint_y=None,
            height=dp(50)
        )

        save_btn = Button(
            text="SAVE CHANGES",
            size_hint_y=None,
            height=dp(55),
            background_color=(0.2, 0.7, 0.3, 1)
        )

        layout.add_widget(mode_spinner)
        layout.add_widget(bill_input)
        layout.add_widget(amount_input)
        layout.add_widget(save_btn)

        popup = Popup(
            title="Edit Bill",
            content=layout,
            size_hint=(0.9, 0.6)
        )

        def save_changes(btn):

            try:

                mode = mode_spinner.text
                bill_no = int(bill_input.text)
                new_amount = float(amount_input.text)

                taxable, cgst, sgst, gst = calculate_gst(new_amount)

                wb = load_workbook(FILE_NAME)
                sheet = wb.active

                found = False

                for row in sheet.iter_rows(min_row=2):

                    if (
                        str(row[1].value).lower() == mode.lower()
                        and row[2].value == bill_no
                    ):

                        row[3].value = new_amount
                        row[4].value = taxable
                        row[5].value = cgst
                        row[6].value = sgst
                        row[7].value = gst

                        found = True
                        break

                wb.save(FILE_NAME)

                if found:
                    self.show_popup(
                        "Success",
                        "Bill updated successfully."
                    )
                else:
                    self.show_popup(
                        "Not Found",
                        "Bill not found."
                    )

                popup.dismiss()

            except Exception as e:
                self.show_popup("Error", str(e))

        save_btn.bind(on_press=save_changes)

        popup.open()


GSTBillingApp().run()