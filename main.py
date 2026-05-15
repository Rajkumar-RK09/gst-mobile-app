# =========================================================
# GST BILLING SYSTEM - MODERN DARK UI
# =========================================================

from kivy.app import App

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput

from kivy.graphics import Color, RoundedRectangle

from kivy.core.window import Window

from openpyxl import Workbook, load_workbook

from datetime import datetime, timedelta

import os

# =========================================================
# MOBILE WINDOW SIZE
# =========================================================
Window.size = (420, 860)

# =========================================================
# SETTINGS
# =========================================================
GST_RATE = 0.05
EXCEL_FILE = "GST_Billing.xlsx"

# =========================================================
# CREATE EXCEL FILE
# =========================================================
if not os.path.exists(EXCEL_FILE):

    wb = Workbook()

    ws = wb.active

    ws.title = "Billing_Data"

    ws.append([
        "Date",
        "Payment Mode",
        "Bill No",
        "Sale Amount",
        "Taxable Amount",
        "CGST",
        "SGST",
        "Total GST"
    ])

    wb.save(EXCEL_FILE)

# =========================================================
# CUSTOM CARD
# =========================================================
class Card(BoxLayout):

    def __init__(self, bg=(0.05,0.08,0.15,1), **kwargs):

        super().__init__(**kwargs)

        self.bg_color = bg

        with self.canvas.before:

            Color(*self.bg_color)

            self.rect = RoundedRectangle(radius=[25])

        self.bind(pos=self.update_rect)
        self.bind(size=self.update_rect)

    def update_rect(self, *args):

        self.rect.pos = self.pos
        self.rect.size = self.size

# =========================================================
# MAIN APP
# =========================================================
class GSTBillingApp(App):

    # =====================================================
    # BUILD
    # =====================================================
    def build(self):

        self.current_date = ""
        self.phonepe_bill_no = 1
        self.cash_bill_no = 1

        self.total_sale = 0
        self.total_taxable = 0

        root = BoxLayout(
            orientation='vertical'
        )

        with root.canvas.before:

            Color(0.01,0.02,0.05,1)

            self.bg = RoundedRectangle()

        root.bind(
            pos=self.update_bg
        )

        root.bind(
            size=self.update_bg
        )

        # =================================================
        # HEADER
        # =================================================
        header = BoxLayout(
            size_hint=(1,0.10),
            padding=10,
            spacing=10
        )

        menu = Label(
            text="☰",
            font_size=34,
            size_hint=(0.15,1),
            color=(1,1,1,1)
        )

        title = Label(
            text="🛡 GST BILLING SYSTEM",
            font_size=24,
            bold=True,
            color=(1,1,1,1)
        )

        settings = Label(
            text="⚙",
            font_size=30,
            size_hint=(0.15,1),
            color=(1,1,1,1)
        )

        header.add_widget(menu)
        header.add_widget(title)
        header.add_widget(settings)

        root.add_widget(header)

        # =================================================
        # SCROLL VIEW
        # =================================================
        scroll = ScrollView()

        self.main_layout = BoxLayout(
            orientation='vertical',
            spacing=15,
            padding=15,
            size_hint_y=None
        )

        self.main_layout.bind(
            minimum_height=self.main_layout.setter('height')
        )

        scroll.add_widget(self.main_layout)

        root.add_widget(scroll)

        # =================================================
        # START SCREEN
        # =================================================
        self.build_start_screen()

        return root

    # =====================================================
    # UPDATE BACKGROUND
    # =====================================================
    def update_bg(self, instance, value):

        self.bg.pos = instance.pos
        self.bg.size = instance.size

    # =====================================================
    # MODERN BUTTON
    # =====================================================
    def modern_button(self, text, color):

        return Button(

            text=text,

            bold=True,

            font_size=20,

            background_normal='',

            background_color=color,

            size_hint_y=None,

            height=75
        )

    # =====================================================
    # INPUT DESIGN
    # =====================================================
    def modern_input(self, hint):

        return TextInput(

            hint_text=hint,

            multiline=False,

            font_size=22,

            size_hint_y=None,

            height=70,

            background_color=(0.08,0.1,0.18,1),

            foreground_color=(1,1,1,1),

            hint_text_color=(0.6,0.6,0.6,1),

            cursor_color=(1,1,1,1),

            padding=[20,20]
        )

    # =====================================================
    # START SCREEN
    # =====================================================
    def build_start_screen(self):

        self.main_layout.clear_widgets()

        start_card = Card(

            orientation='vertical',

            spacing=15,

            padding=20,

            size_hint_y=None,

            height=520
        )

        title = Label(

            text="START BILLING SESSION",

            font_size=26,

            bold=True,

            color=(1,1,1,1),

            size_hint_y=None,

            height=50
        )

        self.date_input = self.modern_input(
            "Enter Date (DD-MM-YY)"
        )

        self.phonepe_input = self.modern_input(
            "Enter PhonePe Starting Bill"
        )

        self.cash_input = self.modern_input(
            "Enter Cash Starting Bill"
        )

        start_btn = self.modern_button(

            "START BILLING",

            (0.35,0.15,0.85,1)
        )

        start_btn.bind(
            on_press=self.start_billing
        )

        self.start_status = Label(

            text="Ready",

            font_size=18,

            color=(0,1,0,1)
        )

        start_card.add_widget(title)
        start_card.add_widget(self.date_input)
        start_card.add_widget(self.phonepe_input)
        start_card.add_widget(self.cash_input)
        start_card.add_widget(start_btn)
        start_card.add_widget(self.start_status)

        self.main_layout.add_widget(start_card)

    # =====================================================
    # START BILLING
    # =====================================================
    def start_billing(self, instance):

        try:

            datetime.strptime(
                self.date_input.text,
                "%d-%m-%y"
            )

            self.current_date = self.date_input.text

            self.phonepe_bill_no = int(
                self.phonepe_input.text
            )

            self.cash_bill_no = int(
                self.cash_input.text
            )

            self.build_dashboard()

        except:

            self.start_status.text = "❌ Invalid Input"

    # =====================================================
    # DASHBOARD
    # =====================================================
    def build_dashboard(self):

        self.main_layout.clear_widgets()

        # =================================================
        # DATE CARD
        # =================================================
        top = Card(

            orientation='horizontal',

            padding=20,

            spacing=15,

            size_hint_y=None,

            height=170
        )

        left = BoxLayout(
            orientation='vertical'
        )

        left.add_widget(Label(
            text="📅",
            font_size=55,
            color=(0.2,0.5,1,1)
        ))

        center = BoxLayout(
            orientation='vertical'
        )

        center.add_widget(Label(
            text="Current Date",
            font_size=18,
            color=(1,1,1,1)
        ))

        self.date_label = Label(
            text=self.current_date,
            font_size=34,
            bold=True,
            color=(1,1,1,1)
        )

        center.add_widget(self.date_label)

        right = GridLayout(
            cols=2,
            spacing=10
        )

        phone_card = Card(
            orientation='vertical',
            bg=(0.15,0.05,0.3,1)
        )

        self.phone_bill_label = Label(
            text=f"{self.phonepe_bill_no}",
            font_size=34,
            bold=True,
            color=(1,1,1,1)
        )

        phone_card.add_widget(Label(
            text="PhonePe",
            font_size=18,
            color=(0.7,0.3,1,1)
        ))

        phone_card.add_widget(self.phone_bill_label)

        cash_card = Card(
            orientation='vertical',
            bg=(0.05,0.25,0.1,1)
        )

        self.cash_bill_label = Label(
            text=f"{self.cash_bill_no}",
            font_size=34,
            bold=True,
            color=(1,1,1,1)
        )

        cash_card.add_widget(Label(
            text="Cash",
            font_size=18,
            color=(0,1,0.3,1)
        ))

        cash_card.add_widget(self.cash_bill_label)

        right.add_widget(phone_card)
        right.add_widget(cash_card)

        top.add_widget(left)
        top.add_widget(center)
        top.add_widget(right)

        self.main_layout.add_widget(top)

        # =================================================
        # ADD BILL CARD
        # =================================================
        bill_card = Card(

            orientation='vertical',

            padding=20,

            spacing=15,

            size_hint_y=None,

            height=320
        )

        bill_card.add_widget(Label(
            text="🧾 ADD NEW BILL",
            font_size=24,
            bold=True,
            color=(1,1,1,1),
            size_hint_y=None,
            height=40
        ))

        self.amount_input = self.modern_input(
            "Enter Sale Amount"
        )

        bill_card.add_widget(self.amount_input)

        btn1 = self.modern_button(
            "🟣 Add PhonePe Bill",
            (0.35,0.10,0.85,1)
        )

        btn2 = self.modern_button(
            "🟢 Add Cash Bill",
            (0.05,0.7,0.15,1)
        )

        btn1.bind(
            on_press=self.add_phonepe_bill
        )

        btn2.bind(
            on_press=self.add_cash_bill
        )

        bill_card.add_widget(btn1)
        bill_card.add_widget(btn2)

        self.main_layout.add_widget(bill_card)

        # =================================================
        # BILL SUMMARY
        # =================================================
        summary = Card(

            orientation='vertical',

            padding=20,

            spacing=15,

            size_hint_y=None,

            height=220
        )

        summary.add_widget(Label(
            text="📊 BILL SUMMARY",
            font_size=24,
            bold=True,
            color=(0.2,0.5,1,1),
            size_hint_y=None,
            height=40
        ))

        self.bill_summary = Label(

            text="Waiting For Entry",

            font_size=22,

            color=(1,1,1,1)
        )

        summary.add_widget(self.bill_summary)

        self.main_layout.add_widget(summary)

        # =================================================
        # DAY SUMMARY
        # =================================================
        total = Card(

            orientation='horizontal',

            spacing=10,

            padding=15,

            size_hint_y=None,

            height=160
        )

        sale_card = Card(
            orientation='vertical',
            bg=(0.05,0.2,0.08,1)
        )

        sale_card.add_widget(Label(
            text="Total Sale",
            font_size=20,
            color=(0,1,0.3,1)
        ))

        self.total_sale_label = Label(
            text="₹ 0.00",
            font_size=34,
            bold=True,
            color=(0,1,0.3,1)
        )

        sale_card.add_widget(self.total_sale_label)

        tax_card = Card(
            orientation='vertical',
            bg=(0.05,0.1,0.25,1)
        )

        tax_card.add_widget(Label(
            text="Total Taxable",
            font_size=20,
            color=(0.2,0.5,1,1)
        ))

        self.total_taxable_label = Label(
            text="₹ 0.00",
            font_size=34,
            bold=True,
            color=(0.2,0.5,1,1)
        )

        tax_card.add_widget(self.total_taxable_label)

        total.add_widget(sale_card)
        total.add_widget(tax_card)

        self.main_layout.add_widget(total)

        # =================================================
        # ACTION BUTTONS
        # =================================================
        action = GridLayout(

            cols=3,

            spacing=10,

            size_hint_y=None,

            height=90
        )

        btn_view = self.modern_button(
            "📋 View",
            (0.05,0.35,0.9,1)
        )

        btn_finish = self.modern_button(
            "📅 Finish",
            (1,0.55,0,1)
        )

        btn_exit = self.modern_button(
            "❌ Exit",
            (0.9,0.1,0.1,1)
        )

        btn_finish.bind(
            on_press=self.finish_day
        )

        btn_exit.bind(
            on_press=self.stop
        )

        btn_view.bind(
            on_press=self.view_bills
        )

        action.add_widget(btn_view)
        action.add_widget(btn_finish)
        action.add_widget(btn_exit)

        self.main_layout.add_widget(action)

        # =================================================
        # EDIT BILL CARD
        # =================================================
        edit = Card(

            orientation='vertical',

            spacing=15,

            padding=20,

            size_hint_y=None,

            height=320
        )

        edit.add_widget(Label(
            text="✏ EDIT BILL",
            font_size=24,
            bold=True,
            color=(1,1,1,1),
            size_hint_y=None,
            height=40
        ))

        self.edit_bill_input = self.modern_input(
            "Enter Bill Number"
        )

        self.edit_amount_input = self.modern_input(
            "Enter New Sale Amount"
        )

        edit.add_widget(self.edit_bill_input)
        edit.add_widget(self.edit_amount_input)

        edit_btns = GridLayout(
            cols=2,
            spacing=10,
            size_hint_y=None,
            height=80
        )

        btn_ep = self.modern_button(
            "🟣 Update PhonePe",
            (0.35,0.10,0.85,1)
        )

        btn_ec = self.modern_button(
            "🟢 Update Cash",
            (0.05,0.7,0.15,1)
        )

        btn_ep.bind(
            on_press=self.update_phonepe_bill
        )

        btn_ec.bind(
            on_press=self.update_cash_bill
        )

        edit_btns.add_widget(btn_ep)
        edit_btns.add_widget(btn_ec)

        edit.add_widget(edit_btns)

        self.main_layout.add_widget(edit)

        # =================================================
        # STATUS
        # =================================================
        self.status_label = Label(

            text="✅ Ready",

            font_size=18,

            color=(1,1,0,1),

            size_hint_y=None,

            height=50
        )

        self.main_layout.add_widget(self.status_label)

    # =====================================================
    # SAVE BILL
    # =====================================================
    def save_bill(self, mode, bill_no, amount,
                  taxable, cgst, sgst, gst):

        wb = load_workbook(EXCEL_FILE)

        ws = wb.active

        ws.append([
            self.current_date,
            mode,
            bill_no,
            amount,
            taxable,
            cgst,
            sgst,
            gst
        ])

        wb.save(EXCEL_FILE)

    # =====================================================
    # PROCESS BILL
    # =====================================================
    def process_bill(self, mode):

        try:

            amount = float(
                self.amount_input.text
            )

            taxable = amount / (1 + GST_RATE)

            gst = amount - taxable

            cgst = gst / 2
            sgst = gst / 2

            if mode == "PhonePe":

                bill_no = self.phonepe_bill_no
                self.phonepe_bill_no += 1

                self.phone_bill_label.text = (
                    str(self.phonepe_bill_no)
                )

            else:

                bill_no = self.cash_bill_no
                self.cash_bill_no += 1

                self.cash_bill_label.text = (
                    str(self.cash_bill_no)
                )

            self.save_bill(
                mode,
                bill_no,
                round(amount,2),
                round(taxable,2),
                round(cgst,2),
                round(sgst,2),
                round(gst,2)
            )

            self.total_sale += amount
            self.total_taxable += taxable

            self.total_sale_label.text = (
                f"₹ {self.total_sale:.2f}"
            )

            self.total_taxable_label.text = (
                f"₹ {self.total_taxable:.2f}"
            )

            self.bill_summary.text = (

                f"Sale Amount : ₹ {amount:.2f}\n\n"

                f"Taxable : ₹ {taxable:.2f}\n\n"

                f"GST : ₹ {gst:.2f}"
            )

            self.status_label.text = (
                f"✅ {mode} Bill Added"
            )

            self.amount_input.text = ""

        except:

            self.status_label.text = (
                "❌ Invalid Amount"
            )

    # =====================================================
    # BUTTON FUNCTIONS
    # =====================================================
    def add_phonepe_bill(self, instance):

        self.process_bill("PhonePe")

    def add_cash_bill(self, instance):

        self.process_bill("Cash")

    # =====================================================
    # VIEW BILLS
    # =====================================================
    def view_bills(self, instance):

        self.status_label.text = (
            "📋 Bills Saved In Excel"
        )

    # =====================================================
    # UPDATE BILL
    # =====================================================
    def update_bill(self, mode):

        try:

            bill_no = int(
                self.edit_bill_input.text
            )

            new_amount = float(
                self.edit_amount_input.text
            )

            wb = load_workbook(EXCEL_FILE)

            ws = wb.active

            found = False

            for row in ws.iter_rows(min_row=2):

                if (
                    str(row[1].value).lower()
                    ==
                    mode.lower()
                    and
                    row[2].value == bill_no
                ):

                    taxable = (
                        new_amount / (1 + GST_RATE)
                    )

                    gst = new_amount - taxable

                    cgst = gst / 2
                    sgst = gst / 2

                    row[3].value = round(new_amount,2)
                    row[4].value = round(taxable,2)
                    row[5].value = round(cgst,2)
                    row[6].value = round(sgst,2)
                    row[7].value = round(gst,2)

                    found = True

                    break

            wb.save(EXCEL_FILE)

            if found:

                self.status_label.text = (
                    f"✅ {mode} Bill Updated"
                )

            else:

                self.status_label.text = (
                    "❌ Bill Not Found"
                )

        except:

            self.status_label.text = (
                "❌ Invalid Input"
            )

    def update_phonepe_bill(self, instance):

        self.update_bill("PhonePe")

    def update_cash_bill(self, instance):

        self.update_bill("Cash")

    # =====================================================
    # FINISH DAY
    # =====================================================
    def finish_day(self, instance):

        wb = load_workbook(EXCEL_FILE)

        ws = wb.active

        ws.append(["","","","","","","",""])

        wb.save(EXCEL_FILE)

        current = datetime.strptime(
            self.current_date,
            "%d-%m-%y"
        )

        next_day = current + timedelta(days=1)

        self.current_date = next_day.strftime(
            "%d-%m-%y"
        )

        self.date_label.text = self.current_date

        self.total_sale = 0
        self.total_taxable = 0

        self.total_sale_label.text = "₹ 0.00"
        self.total_taxable_label.text = "₹ 0.00"

        self.status_label.text = (
            "📅 Day Finished"
        )

# =========================================================
# RUN APP
# =========================================================
GSTBillingApp().run()