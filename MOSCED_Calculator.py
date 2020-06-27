from kivy.app import App
from kivy.uix.textinput import TextInput
from kivy.properties import ObjectProperty
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.lang import Builder
from kivy.core.window import Window
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.dropdown import DropDown
from kivy.uix.button import Button
import math
import re
import openpyxl

# Window_Size
Window.size = (1200, 750)

# Create a reference MOSCED data list
wb = openpyxl.load_workbook('MOSCED_Data.xlsx')
sheet = wb.get_sheet_by_name('Data')
substance_dict = {}
substance_list = []
for row in sheet['B3':'K134']:
    substance_list.append(row[0].value)
    need_list = []
    for object in row[1:]:
        need_list.append(object.value)
    substance_dict[row[0].value] = need_list


# Filtering the Input (not accepting any characters)
class OutputInput(TextInput):
    def insert_text(self, substring, from_undo=False):
        s = ''
        return super(OutputInput, self).insert_text(s, from_undo=from_undo)


# Filtering the Input (only accepts floats)
class FloatInput(TextInput):
    pat = re.compile('[^0-9]')

    def insert_text(self, substring, from_undo=False):
        pat = self.pat
        if '.' in self.text:
            s = re.sub(pat, '', substring)
        else:
            s = '.'.join([re.sub(pat, '', s) for s in substring.split('.', 1)])
        return super(FloatInput, self).insert_text(s, from_undo=from_undo)


# Window Screen Manager
class Window(ScreenManager):
    pass


# StartUp Screen
class StartUp(Screen):
    pass


# Auto Input Screen
class AutoInput(Screen):
    output = ObjectProperty(None)
    temperature = ObjectProperty(None)

    def __init__(self, **kwargs):
        super(AutoInput, self).__init__(**kwargs)
        self.dropdown_1 = DropDown()
        self.dropdown_1.max_height = 250
        for btn_1_index in substance_list:
            self.btn_1 = Button(text=str(btn_1_index), size_hint_y=None, height=44)
            self.btn_1.bind(on_release=lambda btn: self.dropdown_1.select(btn.text))
            self.dropdown_1.add_widget(self.btn_1)
        self.mainbutton_1 = Button(text='Substance 1', size_hint=(0.5, 0.3), pos_hint={"x": 0, "y": 0.75})
        self.mainbutton_1.bind(on_release=self.dropdown_1.open)
        self.dropdown_1.bind(on_select=lambda instance, a: setattr(self.mainbutton_1, 'text', "Substance 1: " + a))
        self.add_widget(self.mainbutton_1)

        self.dropdown_2 = DropDown()
        self.dropdown_2.max_height = 250
        for btn_2_index in substance_list:
            self.btn_2 = Button(text=str(btn_2_index), size_hint_y=None, height=44)
            self.btn_2.bind(on_release=lambda btn: self.dropdown_2.select(btn.text))
            self.dropdown_2.add_widget(self.btn_2)
        self.mainbutton_2 = Button(text='Substance 2', size_hint=(0.5, 0.3), pos_hint={"x": 0.5, "y": 0.75})
        self.mainbutton_2.bind(on_release=self.dropdown_2.open)
        self.dropdown_2.bind(on_select=lambda instance, a: setattr(self.mainbutton_2, 'text', "Substance 2: " + a))
        self.add_widget(self.mainbutton_2)

        self.submit_button = Button(text="Submit", pos_hint={"x": 0.5, "y": 0}, size_hint=(0.5, 0.2), font_size=20)
        self.submit_button.bind(on_press=lambda x: self.press_submit())
        self.add_widget(self.submit_button)

    def calculate(self):
        substance_1 = self.mainbutton_1.text[13:]
        substance_2 = self.mainbutton_2.text[13:]
        v1 = substance_dict[substance_1][2]
        v2 = substance_dict[substance_2][2]
        lambda1 = substance_dict[substance_1][3]
        lambda2 = substance_dict[substance_2][3]
        tau1 = substance_dict[substance_1][4]
        tau2 = substance_dict[substance_2][4]
        rho1 = substance_dict[substance_1][5]
        rho2 = substance_dict[substance_2][5]
        alpha1 = substance_dict[substance_1][6]
        alpha2 = substance_dict[substance_2][6]
        beta1 = substance_dict[substance_1][7]
        beta2 = substance_dict[substance_2][7]
        R = 8.3144598
        temp = list(map(float, self.temperature.text.split()))
        for T in temp:
            def powerT1(n, T):
                return n * pow((293 / T), 0.8)

            def powerT2(n, T):
                return n * pow((293 / T), 0.4)

            POL = pow(rho1, 4) * (1.15 - 1.15 * math.exp(-0.002337 * pow(powerT2(tau1, T), 3))) + 1
            xi1 = 0.68 * (POL - 1) + pow(3.4 - (2.4 * math.exp(-0.002687 * pow(alpha1 * beta1, 1.5))),
                                         pow((293 / T), 2))
            psi1 = POL + 0.002629 * powerT1(alpha1, T) * powerT1(beta1, T)
            aa = 0.953 - 0.002314 * (pow(powerT2(tau2, T), 2) + powerT1(alpha2, T) * powerT1(beta2, T))
            d12 = math.log(pow(v2 / v1, aa)) + 1 - pow(v2 / v1, aa)
            activity_coefficient = (v2 / (R * T)) * (pow(lambda1 - lambda2, 2) +
                                                     (pow(rho1, 2) * pow(rho2, 2) * pow(
                                                         powerT2(tau1, T) - powerT2(tau2, T),
                                                         2)) / psi1 +
                                                     (powerT1(alpha1, T) - powerT1(alpha2, T)) * (
                                                             powerT1(beta1, T) - powerT1(beta2, T)) / xi1) + d12
        return str(activity_coefficient)

    def show_popup_KeyError(self):
        overflow_error = BoxLayout()
        label = Label(text="At least one substance is not picked")
        overflow_error.add_widget(label)
        popup = Popup(title='Value Error', content=overflow_error, size_hint=(None, None), size=(400, 400))
        popup.open()

    def show_popup_UnboundLocalError(self):
        overflow_error = BoxLayout()
        label = Label(text=" You forgot the temperature")
        overflow_error.add_widget(label)
        popup = Popup(title='Value Error', content=overflow_error, size_hint=(None, None), size=(400, 400))
        popup.open()

    def press_submit(self):
        try:
            self.output.text = self.calculate()
        except KeyError:
            self.show_popup_KeyError()
        except UnboundLocalError:
            self.show_popup_UnboundLocalError()

# Manual Input Screen
class ManualInput(Screen):
    compound_1_v = ObjectProperty(None)
    compound_2_v = ObjectProperty(None)
    compound_1_lambda = ObjectProperty(None)
    compound_2_lambda = ObjectProperty(None)
    compound_1_tau = ObjectProperty(None)
    compound_2_tau = ObjectProperty(None)
    compound_1_rho = ObjectProperty(None)
    compound_2_rho = ObjectProperty(None)
    compound_1_alpha = ObjectProperty(None)
    compound_2_alpha = ObjectProperty(None)
    compound_1_beta = ObjectProperty(None)
    compound_2_beta = ObjectProperty(None)
    temperature = ObjectProperty(None)
    output_1_to_2 = ObjectProperty(None)
    output_2_to_1 = ObjectProperty(None)

    def calculate_1_to_2(self):
        v1 = float(self.compound_1_v.text)
        v2 = float(self.compound_2_v.text)
        lambda1 = float(self.compound_1_lambda.text)
        lambda2 = float(self.compound_2_lambda.text)
        tau1 = float(self.compound_1_tau.text)
        tau2 = float(self.compound_2_tau.text)
        rho1 = float(self.compound_1_rho.text)
        rho2 = float(self.compound_2_rho.text)
        alpha1 = float(self.compound_1_alpha.text)
        alpha2 = float(self.compound_2_alpha.text)
        beta1 = float(self.compound_1_beta.text)
        beta2 = float(self.compound_2_beta.text)
        R = 8.3144598
        temp = list(map(float, self.temperature.text.split()))
        needed_list = []
        for T in temp:
            def powerT1(n, T):
                return n * pow((293 / T), 0.8)

            def powerT2(n, T):
                return n * pow((293 / T), 0.4)

            POL = pow(rho1, 4) * (1.15 - 1.15 * math.exp(-0.002337 * pow(powerT2(tau1, T), 3))) + 1
            xi1 = 0.68 * (POL - 1) + pow(3.4 - (2.4 * math.exp(-0.002687 * pow(alpha1 * beta1, 1.5))),
                                         pow((293 / T), 2))
            psi1 = POL + 0.002629 * powerT1(alpha1, T) * powerT1(beta1, T)
            aa = 0.953 - 0.002314 * (pow(powerT2(tau2, T), 2) + powerT1(alpha2, T) * powerT1(beta2, T))
            d12 = math.log(pow(v2 / v1, aa)) + 1 - pow(v2 / v1, aa)
            activity_coefficient = (v2 / (R * T)) * (pow(lambda1 - lambda2, 2) +
                                                     (pow(rho1, 2) * pow(rho2, 2) * pow(
                                                         powerT2(tau1, T) - powerT2(tau2, T),
                                                         2)) / psi1 +
                                                     (powerT1(alpha1, T) - powerT1(alpha2, T)) * (
                                                             powerT1(beta1, T) - powerT1(beta2, T)) / xi1) + d12
            needed_list.append(str(math.exp(activity_coefficient)))
        return ",".join(needed_list)

    def calculate_2_to_1(self):
        v2 = float(self.compound_1_v.text)
        v1 = float(self.compound_2_v.text)
        lambda2 = float(self.compound_1_lambda.text)
        lambda1 = float(self.compound_2_lambda.text)
        tau2 = float(self.compound_1_tau.text)
        tau1 = float(self.compound_2_tau.text)
        rho2 = float(self.compound_1_rho.text)
        rho1 = float(self.compound_2_rho.text)
        alpha2 = float(self.compound_1_alpha.text)
        alpha1 = float(self.compound_2_alpha.text)
        beta2 = float(self.compound_1_beta.text)
        beta1 = float(self.compound_2_beta.text)
        R = 8.3144598
        temp = list(map(float, self.temperature.text.split()))
        needed_list = []
        for T in temp:
            def powerT1(n, T):
                return n * pow((293 / T), 0.8)

            def powerT2(n, T):
                return n * pow((293 / T), 0.4)

            POL = pow(rho1, 4) * (1.15 - 1.15 * math.exp(-0.002337 * pow(powerT2(tau1, T), 3))) + 1
            xi1 = 0.68 * (POL - 1) + pow(3.4 - (2.4 * math.exp(-0.002687 * pow(alpha1 * beta1, 1.5))),
                                         pow((293 / T), 2))
            psi1 = POL + 0.002629 * powerT1(alpha1, T) * powerT1(beta1, T)
            aa = 0.953 - 0.002314 * (pow(powerT2(tau2, T), 2) + powerT1(alpha2, T) * powerT1(beta2, T))
            d12 = math.log(pow(v2 / v1, aa)) + 1 - pow(v2 / v1, aa)
            activity_coefficient = (v2 / (R * T)) * (pow(lambda1 - lambda2, 2) +
                                                     (pow(rho1, 2) * pow(rho2, 2) * pow(
                                                         powerT2(tau1, T) - powerT2(tau2, T),
                                                         2)) / psi1 +
                                                     (powerT1(alpha1, T) - powerT1(alpha2, T)) * (
                                                             powerT1(beta1, T) - powerT1(beta2, T)) / xi1) + d12
            needed_list.append(str(math.exp(activity_coefficient)))
        return ",".join(needed_list)

    def show_popup_ValueError(self):
        value_error = BoxLayout()
        label = Label(text="Your input not all float/integer \nPlease check it again")
        value_error.add_widget(label)
        popup = Popup(title='Value Error', content=value_error, size_hint=(None, None), size=(400, 400))
        popup.open()

    def show_popup_OverFlowError(self):
        overflow_error = BoxLayout()
        label = Label(text="The result is too big \nPlease try another set of numbers")
        overflow_error.add_widget(label)
        popup = Popup(title='Value Error', content=overflow_error, size_hint=(None, None), size=(400, 400))
        popup.open()

    def pressed_submit(self):
        try:
            self.output_1_to_2.text = ""
            self.output_2_to_1.text = ""
            self.output_1_to_2.text = self.calculate_1_to_2()
            self.output_2_to_1.text = self.calculate_2_to_1()
        except OverflowError:
            self.show_popup_OverFlowError()
        except ValueError:
            self.show_popup_ValueError()

    def pressed_clear(self):
        self.compound_1_v.text = ""
        self.compound_2_v.text = ""
        self.compound_1_lambda.text = ""
        self.compound_2_lambda.text = ""
        self.compound_1_tau.text = ""
        self.compound_2_tau.text = ""
        self.compound_1_rho.text = ""
        self.compound_2_rho.text = ""
        self.compound_1_alpha.text = ""
        self.compound_2_alpha.text = ""
        self.compound_1_beta.text = ""
        self.compound_2_beta.text = ""
        self.temperature.text = ""
        self.output_1_to_2.text = ""
        self.output_2_to_1.text = ""


kv = Builder.load_file("mymaincal.kv")


class MyMainCal(App):
    def build(self):
        return kv


if __name__ == "__main__":
    MyMainCal().run()
