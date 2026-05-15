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
from kivy.utils import platform

from openpyxl import Workbook, load_workbook

from datetime import datetime, timedelta

import os

# ---------------- ANDROID STORAGE PERMISSION ----------------

if platform == "android":
    from android.permissions import request_permissions, Permission

    request_permissions([
        Permission.READ_EXTERNAL_STORAGE,
        Permission.WRITE_EXTERNAL_STORAGE
    ])

# ---------------- THEME ----------------

BG_COLOR = (0.08, 0.08, 0.08, 1)
CARD_COLOR = (0.12, 0.12, 0.12, 1)
BTN_COLOR = (0.15, 0.55, 0.95, 1)
TEXT_COLOR = (1, 1, 1, 1)

Window.clearcolor = BG_COLOR

GST_RATE = 0.05

# ---------------- STORAGE PATH ----------------

if platform == "android":
    from android.storage import primary_external_storage_path

    BASE_PATH = primary_external_storage_path()
    SAVE_FOLDER = os.path.join(BASE_PATH, "Download", "GST Billing")

else:
    SAVE_FOLDER = "GST Billing"

os.makedirs(SAVE_FOLDER, exist_ok=True)

FILE_NAME = os.path.join(
    SAVE_FOLDER,
    "GST_Billing_Data.xlsx"
)

# ---------------- CREATE EXCEL FILE ----------------

if not os.path.exists(FILE_NAME):

    wb = Workbook()
    sheet = wb.active
    sheet.title = "Billing_Data"

    headers = [
        "Date",
        "Payment Mode",
        "Bill No",
        "Sale Amount",
        "Taxable Amount",
        "CGST",
        "SGST",
        "Total GST"
    ]

    sheet.append(headers)

    wb.save(FILE_NAME)

# ---------------- GST CALCULATION ----------------


def calculate_gst(amount):

    taxable = amount / (1 + GST_RATE)
    gst = amount - taxable

    cgst = gst / 2
    sgst = gst / 2

    return (
        round(taxable, 2),
        round(cgst, 2),
        round(sgst, 2),
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

        # ---------------- TITLE ----------------

        title = Label(
            text="GST BILLING SYSTEM",
            font_size="24sp",
            bold=True,
            size_hint_y=None,
            height=dp(60),
            color=TEXT_COLOR
        )

        root.add_widget(title)

        # ---------------- SCROLL ----------------

        scroll = ScrollView()

        layout = GridLayout(
            cols=1,
            spacing=dp(10),
            padding=dp(5),
            size_hint_y=None
        )

        layout.bind(minimum_height=layout.setter('height'))

        scroll.add_widget(layout)

        root.add_widget(scroll)

        # ---------------- DATE ----------------

        self.date_input = TextInput(
            hint_text="Enter Date (DD-MM-YY)",
            multiline=False,
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=CARD_COLOR,
            foreground_color=TEXT_COLOR
        )

        layout.add_widget(self.date_input)

        # ---------------- BILL NUMBERS ----------------

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

        layout.add_widget(self.phonepe_input)

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

        layout.add_widget(self.cash_input)

        # ---------------- PAYMENT MODE ----------------

        self.mode_spinner = Spinner(
            text="PhonePe",
            values=("PhonePe", "Cash"),
            size_hint_y=None,
            height=dp(55),
            font_size="18sp",
            background_color=BTN_COLOR
        )

        layout.add_widget(self.mode_spinner)

        # ---------------- AMOUNT ----------------

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

        layout.add_widget(self.amount_input)

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

        layout.add_widget(add_btn)

        # ---------------- EDIT BUTTON ----------------

        edit_btn = Button(
            text="VIEW & EDIT BILLS",
            size_hint_y=None,
            height=dp(58),
            font_size="20sp",
            bold=True,
            background_color=(0.9, 0.5, 0.1, 1)
        )

        edit_btn.bind(on_press=self.show_bills)

        layout.add_widget(edit_btn)

        # ---------------- FINISH DAY ----------------

        finish_btn = Button(
            text="FINISH DAY",
            size_hint_y=None,
            height=dp(58),
            font_size="20sp",
            bold=True,
            background_color=(0.2, 0.7, 0.3, 1)
        )

        finish_btn.bind(on_press=self.finish_day)

        layout.add_widget(finish_btn)

        # ---------------- EXIT BUTTON ----------------

        exit_btn = Button(
            text="SAVE & EXIT",
            size_hint_y=None,
            height=dp(58),
            font_size="20sp",
            bold=True,
            background_color=(0.85, 0.2, 0.2, 1)
        )

        exit_btn.bind(on_press=self.exit_app)

        layout.add_widget(exit_btn)

        # ---------------- SUMMARY ----------------

        self.summary_label = Label(
            text="Daily Totals Will Appear Here",
            size_hint_y=None,
            height=dp(220),
            font_size="18sp",
            halign="left",
            valign="top",
            color=TEXT_COLOR
        )

        self.summary_label.bind(
            size=self.summary_label.setter('text_size')
        )

        layout.add_widget(self.summary_label)

        return root

    # ---------------- POPUP ----------------

    def show_popup(self, title, message):

        box = BoxLayout(
            orientation="vertical",
            padding=dp(10),
            spacing=dp(10)
        )

        msg = Label(
            text=message,
            color=TEXT_COLOR
        )

        btn = Button(
            text="OK",
            size_hint_y=None,
            height=dp(50),
            background_color=BTN_COLOR
        )

        box.add_widget(msg)
        box.add_widget(btn)

        popup = Popup(
            title=title,
            content=box,
            size_hint=(0.85, 0.45)
        )

        btn.bind(on_press=popup.dismiss)

        popup.open()

    # ---------------- ADD BILL ----------------

    def add_bill(self, instance):

        try:

            if self.current_date == "":

                self.current_date = self.date_input.text.strip()

                datetime.strptime(
                    self.current_date,
                    "%d-%m-%y"
                )

                self.phonepe_bill_no = int(
                    self.phonepe_input.text
                )

                self.cash_bill_no = int(
                    self.cash_input.text
                )

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
                f"Taxable Amount: ₹ {taxable:.2f}\n"
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

            sheet.append([""])

            wb.save(FILE_NAME)

            self.show_popup(
                "Day Finished",
                f"Date: {self.current_date}\n\n"
                f"Total Sale: ₹ {self.total_sale:.2f}\n"
                f"Total Taxable: ₹ {self.total_taxable:.2f}"
            )

            old_date = datetime.strptime(
                self.current_date,
                "%d-%m-%y"
            )

            new_date = old_date + timedelta(days=1)

            self.current_date = new_date.strftime("%d-%m-%y")

            self.date_input.text = self.current_date

            self.total_sale = 0
            self.total_taxable = 0

        except Exception as e:
            self.show_popup("Error", str(e))

    # ---------------- SHOW BILLS ----------------

    def show_bills(self, instance):

        try:

            wb = load_workbook(FILE_NAME)
            sheet = wb.active

            content = BoxLayout(
                orientation="vertical",
                spacing=dp(10),
                padding=dp(10)
            )

            scroll = ScrollView(size_hint=(1, 1))

            bills_layout = GridLayout(
                cols=1,
                spacing=dp(10),
                size_hint_y=None
            )

            bills_layout.bind(
                minimum_height=bills_layout.setter('height')
            )

            for row in sheet.iter_rows(min_row=2, values_only=True):

                if row[0] is None:
                    continue

                bill_text = (
                    f"{row[0]} | "
                    f"{row[1]} | "
                    f"Bill {row[2]} | "
                    f"₹ {row[3]}"
                )

                lbl = Label(
                    text=bill_text,
                    size_hint_y=None,
                    height=dp(40),
                    font_size="16sp",
                    color=TEXT_COLOR,
                    halign="left",
                    valign="middle"
                )

                lbl.bind(size=lbl.setter('text_size'))

                bills_layout.add_widget(lbl)

            scroll.add_widget(bills_layout)

            content.add_widget(scroll)

            # ---------------- EDIT SECTION ----------------

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

            content.add_widget(mode_spinner)
            content.add_widget(bill_input)
            content.add_widget(amount_input)
            content.add_widget(save_btn)

            popup = Popup(
                title="View & Edit Bills",
                content=content,
                size_hint=(0.95, 0.95)
            )

            def save_changes(btn):

                try:

                    mode = mode_spinner.text
                    bill_no = int(bill_input.text)
                    new_amount = float(amount_input.text)

                    taxable, cgst, sgst, gst = calculate_gst(
                        new_amount
                    )

                    found = False

                    for row in sheet.iter_rows(min_row=2):

                        if (
                            str(row[1].value).lower()
                            == mode.lower()
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
                            "Error",
                            "Bill not found."
                        )

                    popup.dismiss()

                except Exception as e:
                    self.show_popup("Error", str(e))

            save_btn.bind(on_press=save_changes)

            popup.open()

        except Exception as e:
            self.show_popup("Error", str(e))

    # ---------------- EXIT APP ----------------

    def exit_app(self, instance):

        try:

            self.show_popup(
                "Saved",
                f"Excel File Saved Successfully\n\n"
                f"Location:\n{FILE_NAME}"
            )

            App.get_running_app().stop()

        except Exception as e:
            self.show_popup("Error", str(e))


GSTBillingApp().run()