from kivy.clock import Clock
from kivy.uix.behaviors import ButtonBehavior
from kivymd.app import MDApp
from kivymd.uix.anchorlayout import MDAnchorLayout
from kivymd.uix.behaviors import CommonElevationBehavior
from kivymd.uix.boxlayout import MDBoxLayout
from flash_python_file import Flash

class CustomMDAnchorLayout(ButtonBehavior, CommonElevationBehavior, MDAnchorLayout):
    pass
class Interface(MDBoxLayout):
    flashstate=0
    milli_seconds=0
    seconds=0
    hours=0
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        Flash.initialize()
    def increment_seconds(self, *args):
        Interface.milli_seconds+=10
        if(Interface.milli_seconds==100):
            Interface.milli_seconds=0
            Interface.seconds+=1
            if(Interface.seconds==60):
                Interface.seconds=0
                Interface.hours+=1
        self.ids.timeplaceholder.text="{0:0=2d}".format(Interface.hours)+" : "+ "{0:0=2d}".format(Interface.seconds)+" : "+"{0:0=2d}".format(Interface.milli_seconds)
    def flash_on(self):
        if(Interface.flashstate==0):
            print("Flash Light ON")
            Clock.schedule_interval(self.increment_seconds, 1 / 10)
            self.ids.flash_icon.icon="alarm-light-off"
            Interface.flashstate=1
            self.ids.progressbar.start()
            Flash.on()
        else:
            Clock.unschedule(self.increment_seconds)
            Interface.milli_seconds=0
            Interface.seconds=0
            Interface.hours=0
            Interface.flashstate = 0
            print("Flash Light Off")
            self.ids.flash_icon.icon = "alarm-light"
            self.ids.progressbar.stop()
            Flash.off()
class TestApp(MDApp):
    def change_style(self, appbar):
        if(self.theme_cls.theme_style=="Light"):
            self.theme_cls.theme_style="Dark"
            self.theme_cls.primary_palette = "Amber"
            self.theme_cls.accent_hue = "50"
            appbar.right_action_items= [["weather-sunny", lambda x: self.change_style(self)]]

        else:
            self.theme_cls.theme_style="Light"
            appbar.right_action_items = [["weather-night", lambda x: self.change_style(self)]]
            self.theme_cls.accent_hue = "A700"
            self.theme_cls.primary_palette = "Purple"
    def build(self):
        self.theme_cls.primary_palette="Purple"
        self.theme_cls.accent_palette = "Gray"
        self.theme_cls.accent_hue="A700"

TestApp().run()
