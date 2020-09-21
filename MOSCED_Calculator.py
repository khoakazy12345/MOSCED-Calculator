from kivy.app import App
from kivy.uix.textinput import TextInput
from kivy.properties import ObjectProperty, NumericProperty
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.lang import Builder
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.dropdown import DropDown
from kivy.uix.button import Button
from kivy.uix.widget import Widget
import math
import re
import openpyxl

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


# Main Calculation Algorithm
def calculate(v1, v2, lambda1, lambda2, tau1, tau2, rho1, rho2, alpha1, alpha2, beta1, beta2, temp):
    R = 8.3144598
    for T in temp:
        def powerT1(n, T):
            return n * pow((293 / T), 0.8)

        def powerT2(n, T):
            return n * pow((293 / T), 0.4)

        POL = pow(rho1, 4) * (1.15 - 1.15 * math.exp(-0.002337 * pow(powerT2(tau1, T), 3))) + 1
        xi1 = 0.68 * (POL - 1) + pow(3.4 - (2.4 * math.exp(-0.002687 * pow(alpha1 * beta1, 1.5))), pow((293 / T), 2))
        psi1 = POL + 0.002629 * powerT1(alpha1, T) * powerT1(beta1, T)
        aa = 0.953 - 0.002314 * (pow(powerT2(tau2, T), 2) + powerT1(alpha2, T) * powerT1(beta2, T))
        d12 = math.log(pow(v2 / v1, aa)) + 1 - pow(v2 / v1, aa)
        activity_coefficient = (v2 / (R * T)) * (pow(lambda1 - lambda2, 2) +
                                                 (pow(rho1, 2) * pow(rho2, 2) * pow(powerT2(tau1, T) - powerT2(tau2, T),
                                                                                    2)) / psi1 +
                                                 (powerT1(alpha1, T) - powerT1(alpha2, T)) * (
                                                         powerT1(beta1, T) - powerT1(beta2, T)) / xi1) + d12
        return str(math.exp(activity_coefficient))
        

def show_popup_KeyError():
    key_error = BoxLayout()
    label = Label(text="Please input compound parameters or pick from the dropdown(s)")
    key_error.add_widget(label)
    popup = Popup(title='Key Error', content=key_error, size_hint=(None, None), size=(600, 400))
    popup.open()


def show_popup_AttributeError():
    overflow_error = BoxLayout()
    label = Label(text="You forgot the temperature")
    overflow_error.add_widget(label)
    popup = Popup(title='Attribute Error', content=overflow_error, size_hint=(None, None), size=(400, 400))
    popup.open()


def show_popup_OverFlowError():
    overflow_error = BoxLayout()
    label = Label(text="The result is too big \nPlease try another set of numbers")
    overflow_error.add_widget(label)
    popup = Popup(title='Overflow Error', content=overflow_error, size_hint=(None, None), size=(400, 400))
    popup.open()


def show_popup_ValueError():
    value_error = BoxLayout()
    label = Label(text="Compound inputs and/or temperature are not all float/integer \nPlease check it again")
    value_error.add_widget(label)
    popup = Popup(title='Value Error', content=value_error, size_hint=(None, None), size=(600, 400))
    popup.open()


def show_popup_ZeroDivisionError():
    zerodivision_error = BoxLayout()
    label = Label(text="Temperature cannot be zero")
    zerodivision_error.add_widget(label)
    popup = Popup(title='Zero Division Error', content=zerodivision_error, size_hint=(None, None), size=(600, 400))
    popup.open()


def show_popup_VError():
    v_error = BoxLayout()
    label = Label(text="V cannot be zero")
    v_error.add_widget(label)
    popup = Popup(title='V Error', content=v_error, size_hint=(None, None), size=(600, 400))
    popup.open()


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
    output_1 = ObjectProperty(None)
    output_2 = ObjectProperty(None)
    temperature = ObjectProperty(None)

    def __init__(self, **kwargs):
        super(AutoInput, self).__init__(**kwargs)
        self.dropdown_1 = DropDown()
        self.dropdown_1.max_height = 250
        for btn_1_index in substance_list:
            self.btn_1 = Button(text=str(btn_1_index), size_hint_y=None, height=44)
            self.btn_1.bind(on_release=lambda btn: self.dropdown_1.select(btn.text))
            self.dropdown_1.add_widget(self.btn_1)

        self.dropdown_2 = DropDown()
        self.dropdown_2.max_height = 250
        for btn_2_index in substance_list:
            self.btn_2 = Button(text=str(btn_2_index), size_hint_y=None, height=44)
            self.btn_2.bind(on_release=lambda btn: self.dropdown_2.select(btn.text))
            self.dropdown_2.add_widget(self.btn_2)

    def setup_btn_1(self):
        self.ids.btn_1.bind(on_release=self.dropdown_1.open)
        self.dropdown_1.bind(on_select=lambda instance, a: setattr(self.ids.btn_1, 'text', "Compound 1: " + a))

    def setup_btn_2(self):
        self.ids.btn_2.bind(on_release=self.dropdown_2.open)
        self.dropdown_2.bind(on_select=lambda instance, a: setattr(self.ids.btn_2, 'text', "Compound 2: " + a))

    def show_popup_KeyError(self):
        overflow_error = BoxLayout()
        label = Label(text="At least one compound is not picked")
        overflow_error.add_widget(label)
        popup = Popup(title='Overflow Error', content=overflow_error, size_hint=(None, None), size=(400, 400))
        popup.open()

    def press_clear(self):
        self.ids.btn_1.text = "Compound 1: Choose from the dropdown"
        self.ids.btn_2.text = "Compound 2: Choose from the dropdown"
        self.output_1.text = ""
        self.output_2.text = ""
        self.temperature.text = ""

    def press_submit(self):
        substance_1 = self.ids.btn_1.text[12:]
        substance_2 = self.ids.btn_2.text[12:]
        try:
            self.output_1.text = calculate(substance_dict[substance_1][2], substance_dict[substance_2][2],
                                           substance_dict[substance_1][3], substance_dict[substance_2][3],
                                           substance_dict[substance_1][4], substance_dict[substance_2][4],
                                           substance_dict[substance_1][5], substance_dict[substance_2][5],
                                           substance_dict[substance_1][6], substance_dict[substance_2][6],
                                           substance_dict[substance_1][7], substance_dict[substance_2][7],
                                           list(map(float, self.temperature.text.split())))
            self.output_2.text = calculate(substance_dict[substance_2][2], substance_dict[substance_1][2],
                                           substance_dict[substance_2][3], substance_dict[substance_1][3],
                                           substance_dict[substance_2][4], substance_dict[substance_1][4],
                                           substance_dict[substance_2][5], substance_dict[substance_1][5],
                                           substance_dict[substance_2][6], substance_dict[substance_1][6],
                                           substance_dict[substance_2][7], substance_dict[substance_1][7],
                                           list(map(float, self.temperature.text.split())))
        except KeyError:
            show_popup_KeyError()
        except AttributeError:
            show_popup_AttributeError()
        except ValueError:
            show_popup_ValueError()
        except ZeroDivisionError:
            show_popup_ZeroDivisionError()


# Mixed Input Screen
class MixedInput(Screen, Widget):
    v_1 = ObjectProperty(None)
    v_2 = ObjectProperty(None)
    lambda_1 = ObjectProperty(None)
    lambda_2 = ObjectProperty(None)
    tau_1 = ObjectProperty(None)
    tau_2 = ObjectProperty(None)
    rho_1 = ObjectProperty(None)
    rho_2 = ObjectProperty(None)
    alpha_1 = ObjectProperty(None)
    alpha_2 = ObjectProperty(None)
    beta_1 = ObjectProperty(None)
    beta_2 = ObjectProperty(None)
    temperature = ObjectProperty(None)
    output_1_to_2 = ObjectProperty(None)
    output_2_to_1 = ObjectProperty(None)
    font_size = NumericProperty()

    def __init__(self, **kwargs):
        super(MixedInput, self).__init__(**kwargs)
        self.dropdown_1 = DropDown()
        self.dropdown_1.max_height = 350
        for btn_1_index in substance_list:
            self.btn_1 = Button(text=str(btn_1_index), size_hint_y=None, height=44)
            self.btn_1.bind(on_release=lambda btn: self.dropdown_1.select(btn.text))
            self.btn_1.bind(on_release=lambda btn: self.set_up(btn.text))
            self.dropdown_1.add_widget(self.btn_1)

    def add_dropdown(self):
        self.ids.btn.bind(on_release=self.dropdown_1.open)
        self.dropdown_1.bind(on_select=lambda instance, a: setattr(self.ids.btn, 'text', "Compound 1: " + a))

    def set_up(self, substance_1):
        self.v_1.text = str(substance_dict[substance_1][2])
        self.lambda_1.text = str(substance_dict[substance_1][3])
        self.tau_1.text = str(substance_dict[substance_1][4])
        self.rho_1.text = str(substance_dict[substance_1][5])
        self.alpha_1.text = str(substance_dict[substance_1][6])
        self.beta_1.text = str(substance_dict[substance_1][7])

    def press_clear(self):
        self.v_1.text = ""
        self.lambda_1.text = ""
        self.tau_1.text = ""
        self.rho_1.text = ""
        self.alpha_1.text = ""
        self.beta_1.text = ""
        self.v_2.text = ""
        self.lambda_2.text = ""
        self.tau_2.text = ""
        self.rho_2.text = ""
        self.alpha_2.text = ""
        self.beta_2.text = ""
        self.temperature.text = ""
        self.output_1_to_2.text = ""
        self.output_2_to_1.text = ""
        self.ids.btn.text = "Compound 1: Choose one from the list"

    def press_submit(self):
        if self.v_1.text == "0" or self.v_2.text == "0":
            show_popup_VError()
        else:
            substance_1 = self.ids.btn.text[12:]
            try:
                self.output_1_to_2.text = calculate(substance_dict[substance_1][2], float(self.v_2.text),
                                                    substance_dict[substance_1][3], float(self.lambda_2.text),
                                                    substance_dict[substance_1][4], float(self.tau_2.text),
                                                    substance_dict[substance_1][5], float(self.rho_2.text),
                                                    substance_dict[substance_1][6], float(self.alpha_2.text),
                                                    substance_dict[substance_1][7], float(self.beta_2.text),
                                                    list(map(float, self.temperature.text.split())))
                self.output_2_to_1.text = calculate(float(self.v_2.text), substance_dict[substance_1][2],
                                                    float(self.lambda_2.text), substance_dict[substance_1][3],
                                                    float(self.tau_2.text), substance_dict[substance_1][4],
                                                    float(self.rho_2.text), substance_dict[substance_1][5],
                                                    float(self.alpha_2.text), substance_dict[substance_1][6],
                                                    float(self.beta_2.text), substance_dict[substance_1][7],
                                                    list(map(float, self.temperature.text.split())))
            except KeyError:
                show_popup_KeyError()
            except AttributeError:
                show_popup_AttributeError()
            except ValueError:
                show_popup_ValueError()
            except OverflowError:
                show_popup_OverFlowError()


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


    def print(self):
        print("it worked")

    def pressed_submit(self):
        if self.compound_1_v.text == "0" or self.compound_2_v.text == "0":
            show_popup_VError()
        else:
            try:
                self.output_1_to_2.text = ""
                self.output_2_to_1.text = ""
                self.output_1_to_2.text = calculate(float(self.compound_1_v.text), float(self.compound_2_v.text),
                                                float(self.compound_1_lambda.text), float(self.compound_2_lambda.text),
                                                float(self.compound_1_tau.text), float(self.compound_2_tau.text),
                                                float(self.compound_1_rho.text), float(self.compound_2_rho.text),
                                                float(self.compound_1_alpha.text), float(self.compound_2_alpha.text),
                                                float(self.compound_1_beta.text), float(self.compound_2_beta.text),
                                                list(map(float, self.temperature.text.split())))
                self.output_2_to_1.text = calculate(float(self.compound_2_v.text), float(self.compound_1_v.text),
                                                float(self.compound_2_lambda.text), float(self.compound_1_lambda.text),
                                                float(self.compound_2_tau.text), float(self.compound_1_tau.text),
                                                float(self.compound_2_rho.text), float(self.compound_1_rho.text),
                                                float(self.compound_2_alpha.text), float(self.compound_1_alpha.text),
                                                float(self.compound_2_beta.text), float(self.compound_1_beta.text),
                                                list(map(float, self.temperature.text.split())))
            except OverflowError:
                show_popup_OverFlowError()
            except ValueError:
                show_popup_ValueError()
            except ZeroDivisionError:
                show_popup_ZeroDivisionError()

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


class MOSCED_Calculator(App):
    def build(self):
        return kv


if __name__ == "__main__":
    MOSCED_Calculator().run()
