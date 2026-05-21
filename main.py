import os
import sys
import math
import openpyxl

# --- ПУТИ К ФАЙЛАМ (Оптимизировано для Android 13-14) ---
# Мы НЕ используем os.chdir, так как это вызывает вылет на новых Android.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BACKGROUND_PATH = os.path.join(BASE_DIR, 'background.png')
ICON_PATH = os.path.join(BASE_DIR, 'icon.png')
LOGO_PATH = os.path.join(BASE_DIR, 'logo.png')

# Пути к твоим базам данных Excel
EXCEL_PATH = os.path.join(BASE_DIR, 'data', 'nkt_base.xlsx')
CA320_PATH = os.path.join(BASE_DIR, 'data', 'ca320_stats.xlsx')
VZD_PATH = os.path.join(BASE_DIR, 'data', 'vzd_catalog.xlsx')

# --- БАЗОВЫЕ ИМПОРТЫ KIVY ---
from kivy.metrics import dp, sp
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition
from kivy.uix.widget import Widget
from kivy.uix.scrollview import ScrollView
from kivy.uix.floatlayout import FloatLayout
from kivy.core.window import Window
from kivy.core.image import Image as CoreImage
from kivy.graphics import Color, Line, Rectangle, Ellipse
from kivy.clock import Clock
from kivy.properties import ObjectProperty, NumericProperty, StringProperty
from kivy.storage.jsonstore import JsonStore

# --- КОМПОНЕНТЫ KIVYMD ---
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.anchorlayout import MDAnchorLayout
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.textfield import MDTextField
from kivymd.uix.label import MDLabel
from kivymd.uix.fitimage import FitImage
from kivymd.uix.button import (
    MDRaisedButton, 
    MDFillRoundFlatIconButton, 
    MDFillRoundFlatButton, 
    MDIconButton,
    MDFlatButton
)
from kivymd.uix.progressbar import MDProgressBar
from kivymd.uix.snackbar import Snackbar
from kivymd.uix.card import MDCard
from kivymd.uix.dialog import MDDialog

# Обычные Kivy компоненты для верстки
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.core.image import Image as CoreImage
# Фиксация масштаба шрифта, чтобы системные настройки Android не ломали меню
from kivy.metrics import Metrics
Metrics.font_scale = 1.0
import webbrowser

class BaseScreen(MDScreen):   # Фундамент. Отвечает за «умный» фон, который подстраивается под экран без искажений.
    def __init__(self, **kw):
        super().__init__(**kw)
        app = MDApp.get_running_app()
        # Пытаемся получить текстуру фона, загруженную в главном классе приложения
        self.bg_texture = getattr(app, 'global_texture', None)
        
        with self.canvas.before:
            # Слой 1: Сама картинка фона
            Color(1, 1, 1, 1)
            self.bg_rect = Rectangle(pos=self.pos, size=self.size)
            if self.bg_texture:
                self.bg_rect.texture = self.bg_texture
            
            # Слой 2: Затемнение (эффект тонировки)
            # Черный цвет с прозрачностью 70%
            Color(0, 0, 0, 0.7) 
            self.fade_rect = Rectangle(pos=self.pos, size=self.size)
        
        # Обновляем размеры при изменении окна (поворот экрана)
        self.bind(size=self._update_bg, pos=self._update_bg)

    def _update_bg(self, instance, value):
        self.fade_rect.pos = self.bg_rect.pos = instance.pos
        self.fade_rect.size = self.bg_rect.size = instance.size
        
        if self.bg_texture:
            win_w, win_h = instance.size
            img_w, img_h = self.bg_texture.size
            win_ratio = win_w / max(win_h, 1)
            img_ratio = img_w / max(img_h, 1)

            # Алгоритм обрезки лишнего, чтобы сохранить пропорции (Aspect Ratio)
            if win_ratio > img_ratio:
                h_diff = (1 - img_ratio / win_ratio) / 2
                coords = (0, 1 - h_diff, 1, 1 - h_diff, 1, h_diff, 0, h_diff)
            else:
                w_diff = (1 - win_ratio / img_ratio) / 2
                coords = (w_diff, 1, 1 - w_diff, 1, 1 - w_diff, 0, w_diff, 0)
            self.bg_rect.tex_coords = coords

class WellSchematic(Widget):               #Движок отрисовки. Рисует </kivy.uix.image.Image object at 0x00000174CD1C50F0>конструкцию скважины (колонну, хвостовик, ГНО) и анимирует движение раствора.
    def __init__(self, hz=2000, hv=1500, gno_type="Воронка", l_hvost=0, method="ПРЯМАЯ", **kwargs):
        super().__init__(**kwargs)
        try:
            self.hz = float(hz) if hz else 2000.0
            self.hv = float(hv) if hv else 1500.0
            self.l_hvost = float(l_hvost) if l_hvost else 0.0
        except:
            self.hz, self.hv, self.l_hvost = 2000.0, 1500.0, 0.0
            
        self.gno_type = str(gno_type).upper()
        self.method = str(method).upper()
        self.progress = 0
        Clock.schedule_interval(self.animate, 1/60)
        self.bind(pos=self.update_canvas, size=self.update_canvas)

    def animate(self, dt):
        if self.progress < 2.0:
            self.progress += 0.01
            self.update_canvas()
            return True
        return False

    def update_canvas(self, *args):
        self.canvas.clear()
        cx, cy = self.center_x, self.y
        w, h = self.size
        m = dp(90) # Верхний отступ для ФА
        s = (h - m * 2) / self.hz if self.hz > 0 else 1
        
        y_t = cy + h - m
        y_v = y_t - (self.hv * s)
        y_z = cy + m
        y_h = y_t - ((self.hz - self.l_hvost) * s)

        with self.canvas:
            # --- 1. ВОЗВРАЩАЕМ РОДНОЙ ФОН ПРИЛОЖЕНИЯ ИЗ ГЛАВНОГО ЯДРА ---
            from kivymd.app import MDApp
            app = MDApp.get_running_app()
            
            if hasattr(app, 'global_texture') and app.global_texture:
                # Если текстура фона загружена, накатываем её на весь экран схемы
                Color(1, 1, 1, 0.45) # opacity=0.45, как на главном экране
                Rectangle(pos=self.pos, size=self.size, texture=app.global_texture)
            else:
                # На случай, если картинка не нашлась — строгий темный цвет, чтобы не было пустоты
                Color(0.08, 0.09, 0.12, 1) 
                Rectangle(pos=self.pos, size=self.size)

            # --- 2. ВНУТРЕННИЙ ТЕМНЫЙ ФОН СКВАЖИНЫ (СТРОГО ПОД ЗЕМЛЕЙ) ---
            Color(0.1, 0.1, 0.15, 1) 
            Rectangle(pos=(cx - dp(30), y_z - dp(10)), size=(dp(60), y_t - y_z + dp(10)))

            # --- 3. КОЛОННА (Металлические стенки скважины 168мм) ---
            Color(0.4, 0.4, 0.45, 1) # Тень
            Rectangle(pos=(cx - dp(26), y_h), size=(dp(3), y_t - y_h))
            Rectangle(pos=(cx + dp(23), y_h), size=(dp(3), y_t - y_h))
            Color(0.8, 0.8, 0.85, 1) # Блик
            Rectangle(pos=(cx - dp(25), y_h), size=(dp(1.5), y_t - y_h))
            Rectangle(pos=(cx + dp(24), y_h), size=(dp(1.5), y_t - y_h))

            # Хвостовик 114
            if self.l_hvost > 0:
                Color(0.3, 0.3, 0.35, 1)
                Rectangle(pos=(cx - dp(19), y_z), size=(dp(2), y_h - y_z))
                Rectangle(pos=(cx + dp(17), y_z), size=(dp(2), y_h - y_z))
                Line(points=[cx - dp(25), y_h, cx - dp(18), y_h, cx - dp(18), y_z, cx + dp(18), y_z, cx + dp(18), y_h, cx + dp(25), y_h], width=dp(1.5))
            else:
                Line(points=[cx - dp(25), y_z, cx + dp(25), y_z], width=dp(2)) # Забой

            # --- 4. АНИМАЦИЯ РАСТВОРА (Твоя оригинальная логика) ---
            p = self.progress
            
            # ЭТАП 1: Заполнение НКТ
            p1 = min(p / 0.6, 1.0)
            h1 = p1 * (self.hv * s)
            Color(0.1, 0.7, 0.2, 0.9) 
            Rectangle(pos=(cx - dp(6), y_t - h1), size=(dp(12), h1))
            Color(1, 1, 1, 0.2) 
            Rectangle(pos=(cx - dp(1), y_t - h1), size=(dp(2), h1))

            # ЭТАП 2: Забойное пространство
            if p > 0.6:
                p2 = min((p - 0.6) / 0.6, 1.0)
                h_total_pod = (self.hz - self.hv) * s
                fill_h2 = p2 * h_total_pod
                
                Color(0.05, 0.5, 0.1, 0.8) 
                if y_v > y_h:
                    h_col = min(fill_h2, (y_v - y_h))
                    Rectangle(pos=(cx - dp(23), y_v - h_col), size=(dp(46), h_col))
                    if fill_h2 > (y_v - y_h):
                        h_z = min(fill_h2 - (y_v - y_h), (y_h - y_z))
                        Rectangle(pos=(cx - dp(17), y_h - h_z), size=(dp(34), h_z))
                else:
                    h_z = min(fill_h2, (y_v - y_z))
                    Rectangle(pos=(cx - dp(17), y_v - h_z), size=(dp(34), h_z))

            # ЭТАП 3: Подъем по затрубу
            if p > 1.2 and "УСТ." not in self.gno_type:
                p3 = min((p - 1.2) / 0.8, 1.0)
                h3 = p3 * (self.hv * s)
                Color(0.1, 0.6, 0.2, 0.8)
                Rectangle(pos=(cx - dp(23), y_v), size=(dp(17), h3))
                Rectangle(pos=(cx + dp(6), y_v), size=(dp(17), h3))

            # --- 5. НКТ ---
            Color(0.9, 0.9, 0.9, 0.6)
            Line(points=[cx - dp(7), y_t, cx - dp(7), y_v], width=dp(1))
            Line(points=[cx + dp(7), y_t, cx + dp(7), y_v], width=dp(1))

            # --- 6. ОБОРУДОВАНИЕ (ГНО) ---
            if "ПАКЕР" in self.gno_type:
                Color(0.2, 0.2, 0.2, 1) 
                Rectangle(pos=(cx - dp(24), y_v), size=(dp(17), dp(16)))
                Rectangle(pos=(cx + dp(7), y_v), size=(dp(17), dp(16)))
                Color(1, 0.7, 0, 1) 
                Rectangle(pos=(cx - dp(24), y_v + dp(6)), size=(dp(17), dp(3)))
                Rectangle(pos=(cx + dp(7), y_v + dp(6)), size=(dp(17), dp(3)))
            
            elif "ЭЦН" in self.gno_type:
                Color(0.6, 0.6, 0.65, 1)
                Rectangle(pos=(cx - dp(6), y_v - dp(50)), size=(dp(12), dp(50)))
                Color(0.8, 0.1, 0.1, 1)
                Line(points=[cx + dp(6), y_t, cx + dp(6), y_v - dp(25)], width=dp(1))
                Color(0, 0, 0, 0.5)
                for i in range(5):
                    Line(points=[cx - dp(6), y_v - dp(35 + i * 3), cx + dp(6), y_v - dp(35 + i * 3)], width=dp(1))
            
            else: 
                Color(1, 0.8, 0, 1)
                Line(points=[cx - dp(7), y_v, cx - dp(16), y_v - dp(12)], width=dp(1.5))
                Line(points=[cx + dp(7), y_v, cx + dp(16), y_v - dp(12)], width=dp(1.5))
                
            # --- 7. ЗОНА ПЕРФОРАЦИИ: ТОЧКАМИ ПО СПИРАЛИ ---
            Color(1, 1, 1, 0.6) 
            perf_start_y = y_z + dp(15)
            Ellipse(pos=(cx - dp(20), perf_start_y), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx - dp(10), perf_start_y + dp(10)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx, perf_start_y + dp(20)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx + dp(10), perf_start_y + dp(30)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx + dp(17), perf_start_y + dp(40)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx - dp(15), perf_start_y + dp(50)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx - dp(5), perf_start_y + dp(60)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx + dp(5), perf_start_y + dp(70)), size=(dp(3.5), dp(3.5)))
            Ellipse(pos=(cx + dp(15), perf_start_y + dp(80)), size=(dp(3.5), dp(3.5)))

            # =========================================================================
            # --- 8. ФОНТАННАЯ АРМАТУРА (ЧИСТО СИДИТ НА НАШЕМ РОДНОМ ФОНЕ ПРИЛОЖЕНИЯ) ---
            # =========================================================================
            ground_y = y_t - dp(5)
            Color(0.4, 0.4, 0.45, 1)
            Line(points=[cx - dp(100), ground_y, cx + dp(100), ground_y], width=dp(2))

            Color(0.15, 0.45, 0.75, 1) 
            Rectangle(pos=(cx - dp(32), ground_y), size=(dp(64), dp(8)))
            Rectangle(pos=(cx - dp(16), ground_y + dp(8)), size=(dp(32), dp(16)))
            Rectangle(pos=(cx - dp(52), ground_y + dp(10)), size=(dp(36), dp(12)))
            Rectangle(pos=(cx + dp(16), ground_y + dp(10)), size=(dp(36), dp(12)))
            Rectangle(pos=(cx - dp(14), ground_y + dp(24)), size=(dp(28), dp(16)))
            Rectangle(pos=(cx - dp(14), ground_y + dp(40)), size=(dp(28), dp(20)))
            Rectangle(pos=(cx - dp(56), ground_y + dp(44)), size=(dp(42), dp(12)))
            Rectangle(pos=(cx + dp(14), ground_y + dp(44)), size=(dp(42), dp(12)))
            Rectangle(pos=(cx - dp(12), ground_y + dp(60)), size=(dp(24), dp(14)))

            # Манометры
            Color(1, 1, 1, 1)
            Ellipse(pos=(cx - dp(72), ground_y + dp(43)), size=(dp(14), dp(14))) 
            Ellipse(pos=(cx - dp(7), ground_y + dp(74)), size=(dp(14), dp(14)))  
            Color(0.8, 0.1, 0.1, 1)
            Line(points=[cx - dp(65), ground_y + dp(50), cx - dp(69), ground_y + dp(46)], width=dp(1.5))
            Line(points=[cx, ground_y + dp(74), cx + dp(2), ground_y + dp(84)], width=dp(1.5))

class WellViewScreen(BaseScreen):   # Экран визуализации. «Сцена», на которой запускается анимация закачки раствора.
    def __init__(self, **kw):
        super().__init__(**kw)
        # 1. Создаем основной макет
        self.layout = MDBoxLayout(orientation='vertical')
        
        # 2. Шапка (Твой стандарт)
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(45),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        from kivymd.uix.button import MDIconButton
        btn_back = MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: self.go_back()
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(
            text="СИМУЛЯЦИЯ ГЛУШЕНИЯ", theme_text_color="Custom", 
            text_color=(1,1,1,1), font_style="H6"
        ))
        self.layout.add_widget(self.toolbar)

        # 3. ТА САМАЯ ОБЛАСТЬ, где будет рисоваться скважина
        self.draw_area = MDBoxLayout()
        self.layout.add_widget(self.draw_area)
        
        # 4. Кнопка повтора (Зеленая, как на производстве)
        btn_repeat = MDFillRoundFlatButton(
            text="ПОВТОРИТЬ ЗАКАЧКУ", size_hint_x=1, height=dp(50),
            md_bg_color=(0.1, 0.6, 0.2, 1),
            on_release=lambda x: self.on_enter() # Перезапуск метода отрисовки
        )
        self.layout.add_widget(btn_repeat)
        
        self.add_widget(self.layout)

    def on_enter(self):
        """Срабатывает каждый раз при переходе на этот экран"""
        self.draw_area.clear_widgets()
        try:
            # Получаем доступ к данным из основного калькулятора
            calc = self.manager.get_screen('calc')
            
            # Собираем данные из полей ввода
            hz_val = float(calc.f_ckod.text or 2000)
            lh_val = float(calc.f_l_hvost.text or 0)
            
            # Логика глубины воронки/насоса
            if "ЭЦН" in calc.selected_gno.upper():
                # Для ЭЦН глубина — это сумма секций НКТ
                hv_val = float(calc.f_l73.text or 0) + float(calc.f_l60.text or 0)
            else:
                hv_val = float(calc.f_h_v.text or 1500)
            
            # Создаем динамический виджет схемы
            schematic = WellSchematic(
                hz=hz_val, 
                hv=hv_val, 
                gno_type=calc.selected_gno, 
                l_hvost=lh_val, 
                # Если пакер снят или просто воронка — метод считается обраткой/прямой
                method="ОБРАТКА" if ("ВОРОНКА" in calc.selected_gno.upper() or "СНЯТ" in calc.selected_gno.upper()) else "ПРЯМАЯ"
            )
            # Вставляем схему в область отрисовки
            self.draw_area.add_widget(schematic)
        except Exception as e:
            print(f"Ошибка при инициализации симуляции: {e}")

    def on_leave(self):
        """Очищаем память при уходе с экрана"""
        self.draw_area.clear_widgets()

    def go_back(self):
        self.manager.current = 'kill_res'

class LoadingScreen(MDScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        
        # Строгий чистый черный фон без картинок
        with self.canvas.before:
            Color(0, 0, 0, 1) 
            self.bg_rect = Rectangle(pos=self.pos, size=self.size)
            
        self.bind(size=self._update_bg, pos=self._update_bg)

        layout = MDBoxLayout(orientation='vertical', padding=dp(50))
        layout.add_widget(MDBoxLayout(size_hint_y=0.45)) 

        # Тонкая линия загрузки (ниточка)
        self.progress = MDProgressBar(
            value=0, max=100, color=(0.1, 0.4, 0.8, 1), 
            size_hint_x=0.5, size_hint_y=None, height=dp(2),           
            pos_hint={'center_x': 0.5}
        )
        layout.add_widget(self.progress)

        # Твоё название приложения OIL MATE вместо фамилии
        layout.add_widget(MDLabel(
            text="OIL MATE", halign="center", font_style="Button",
            theme_text_color="Custom", text_color=(1, 1, 1, 0.6),
            size_hint_y=None, height=dp(40)
        ))
        
        layout.add_widget(MDBoxLayout(size_hint_y=0.45)) 
        self.add_widget(layout)

    def _update_bg(self, instance, value):
        self.bg_rect.pos = instance.pos
        self.bg_rect.size = instance.size

    def on_enter(self):
        self.progress.value = 0
        # Никаких пауз, запускаем прогресс-бар сразу
        Clock.schedule_interval(self.update_progress, 0.02)

    def update_progress(self, dt):
        if self.progress.value < 100:
            self.progress.value += 2
            return True
        else:
            # Переход в меню
            self.manager.current = 'menu'
            app = MDApp.get_running_app()
            Clock.schedule_once(lambda x: app.show_disclaimer(), 0.3)
            return False

class CalcScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.selected_gno = "Воронка"
        self.tech_nkt_name = ""  # Текстовое имя выбранной НКТ
        self.v_pipe = None       # Внутренний объем трубы для глушения (л/м)
        self.vtulka, self.gear, self.rpm = "115", "3", 1000
        
        # Категория опасности (по умолчанию 3-я) с критериями
        self.well_category = "3 КАТЕГОРИЯ"
        
        # Метод глушения по умолчанию и словарь человеческих названий
        self.kill_method = "STAND"
        self.method_names = {
            "STAND": "ШТАТНЫЙ РЕЖИМ ГЛУШЕНИЯ",
            "DRILLER": "МЕТОД БУРИЛЬЩИКА (2 ЦИКЛА)",
            "WAIT": "ОЖИДАНИЕ И УТЯЖЕЛЕНИЕ (1 ЦИКЛ)",
            "BULL": "ЗАДАВКА В ПЛАСТ "
        }

        self.q_coeffs = {
            "100": {"2": 0.00171, "3": 0.00342, "4": 0.00513, "5": 0.00777},
            "115": {"2": 0.00227, "3": 0.00453, "4": 0.00680, "5": 0.01020},
            "127": {"2": 0.00277, "3": 0.00555, "4": 0.00832, "5": 0.01247}
        }

        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))
        
        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(45), 
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        btn_back = MDIconButton(
            icon="arrow-left", pos_hint={'center_y': .5}, 
            theme_text_color="Custom", text_color=(1, 1, 1, 1), 
            on_release=lambda x: self.go_back()
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(
            text="ВЫБОР ТЕХНОЛОГИИ ГЛУШЕНИЯ", theme_text_color="Custom", 
            text_color=(1,1,1,1), font_style="H6"
        ))
        main_layout.add_widget(self.toolbar)

        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(10), padding=dp(20))

        # 2. КНОПКА КАТЕГОРИИ ОПАСНОСТИ (ЯРКО-КРАСНАЯ)
        self.container.add_widget(MDLabel(text="КАТЕГОРИЯ СКВАЖИНЫ ПО ПЛАНУ РАБОТ:", font_style="Caption", theme_text_color="Hint"))
        anchor_cat = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(45))
        self.btn_category = MDFillRoundFlatButton(
            text=f"ОПАСНОСТЬ: {self.well_category}", size_hint=(1, 1), 
            md_bg_color=(0.85, 0.15, 0.15, 1), on_release=self.menu_category
        )
        self.btn_category.size_hint_max_x = dp(2000)
        anchor_cat.add_widget(self.btn_category)
        self.container.add_widget(anchor_cat)

        # 3. РЯД КНОПОК ГНО И НКТ
        self.top_btns_row = MDBoxLayout(orientation='horizontal', spacing=dp(5), size_hint_y=None, height=dp(45), size_hint_x=1)
        self.btn_gno = MDFillRoundFlatButton(text=f"ГНО: {self.selected_gno.upper()}", size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), on_release=self.menu_gno)
        self.nkt_box_wrap = MDBoxLayout(size_hint=(1, 1)) 
        self.btn_nkt = MDFillRoundFlatButton(text="ВЫБРАТЬ НКТ", size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), on_release=self.open_nkt_menu)
        self.nkt_box_wrap.add_widget(self.btn_nkt)
        
        self.top_btns_row.add_widget(self.btn_gno)
        self.top_btns_row.add_widget(self.nkt_box_wrap)
        self.container.add_widget(self.top_btns_row)

        # 4. ПОЛЯ ВВОДА (ГЕОМЕТРИЯ)
        self.f_d_col = self.create_small_field("Ø экспл. колонны (мм)", "168")
        self.f_ckod = self.create_small_field("ЦКОД / Иск. забой (м)", "3200")
        self.f_l_hvost = self.create_small_field("Длина хвостовика 114 (м)", "0")
        self.f_h_tvd = self.create_small_field("Глубина по вертикали (м)", "2200")
        
        for f in [self.f_d_col, self.f_ckod, self.f_l_hvost, self.f_h_tvd]:
            self.container.add_widget(f)

        # Динамический блок для полей ЭЦН или Воронки
        self.gno_fields_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(10))
        self.container.add_widget(self.gno_fields_box)

        # 5. ДАННЫЕ ДАВЛЕНИЙ
        self.container.add_widget(MDLabel(text="ПАРАМЕТРЫ ДАВЛЕНИЙ И РАСТВОРА:", font_style="Caption", theme_text_color="Hint"))
        self.f_p_pl = self.create_small_field("P пласт (атм) — ЕСЛИ ИЗВЕСТНО", "")
        self.f_p_ust = self.create_small_field("P трубное (атм)", "0")
        self.f_p_zatr = self.create_small_field("P затрубное (атм)", "0")
        self.f_rho = self.create_small_field("Текущая плотность раствора в скв. (г/см³)", "1.10")
        
        for f in [self.f_p_pl, self.f_p_ust, self.f_p_zatr, self.f_rho]:
            self.container.add_widget(f)

        # 6. ВЫБОР МЕТОДА ГЛУШЕНИЯ
        self.container.add_widget(MDLabel(text="МЕТОД ГЛУШЕНИЯ:", font_style="Caption", theme_text_color="Hint"))
        anchor_method = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(50))
        self.btn_method = MDFillRoundFlatButton(
            text="", size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), on_release=self.menu_kill_methods
        )
        self.btn_method.size_hint_max_x = dp(2000)
        anchor_method.add_widget(self.btn_method)
        self.container.add_widget(anchor_method)

        # 7. СЕКЦИЯ АГРЕГАТА ЦА-320
        self.container.add_widget(MDLabel(text="РЕЖИМ РАБОТЫ ЦА-320", font_style="Subtitle2", bold=True, theme_text_color="Secondary", halign="center"))
        pump_grid = MDGridLayout(cols=3, spacing=dp(5), size_hint_y=None, height=dp(45), size_hint_x=1)
        
        def create_p_btn(text, func):
            wrap = MDBoxLayout(size_hint_x=1)
            btn = MDFillRoundFlatButton(text=text, size_hint=(1,1), md_bg_color=(0.1, 0.4, 0.8, 1), on_release=func)
            wrap.add_widget(btn)
            return btn, wrap

        self.btn_v, w1 = create_p_btn(f"ВТ:{self.vtulka}", self.menu_v)
        self.btn_g, w2 = create_p_btn(f"СКОР:{self.gear}", self.menu_g)
        self.btn_r, w3 = create_p_btn(f"ОБ:{self.rpm}", self.menu_r)
        for w in [w1, w2, w3]: pump_grid.add_widget(w)
        self.container.add_widget(pump_grid)

        # 8. КНОПКА РАСЧЕТА / ИНДИКАТОР ПРЕДУПРЕЖДЕНИЯ (Ширина привязана к якорю)
        self.calc_anchor = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(60))
        self.container.add_widget(self.calc_anchor)

        scroll.add_widget(self.container)
        main_layout.add_widget(scroll)
        self.add_widget(main_layout)
        
        self.refresh_ui_elements()


    def create_small_field(self, hint, val=""):
        f = MDTextField(hint_text=hint, text=val, mode="fill", font_size="15sp", size_hint_y=None, height=dp(34), input_filter='float')
        f.bind(focus=lambda ins, val: Clock.schedule_once(lambda dt: ins.select_all(), 0.1) if val else None)
        return f
    def refresh_ui_elements(self):
        self.gno_fields_box.clear_widgets()
        
        if self.tech_nkt_name:
            self.btn_nkt.text = f"НКТ {self.tech_nkt_name} ({self.v_pipe:.2f} л/м)"
        else:
            self.btn_nkt.text = "ВЫБРАТЬ НКТ"

        self.btn_gno.text = f"ГНО: {self.selected_gno.upper()}"
        self.btn_method.text = self.method_names[self.kill_method]

        if "ЭЦН" in self.selected_gno.upper():
            self.nkt_box_wrap.size_hint_x, self.nkt_box_wrap.opacity, self.nkt_box_wrap.disabled = 0, 0, True
            self.btn_gno.size_hint_x = 1
            
            self.f_l73 = self.create_small_field("Длина НКТ 73 (м)", "1000")
            self.f_l60 = self.create_small_field("Длина НКТ 60 (м)", "500")
            self.gno_fields_box.add_widget(self.f_l73)
            self.gno_fields_box.add_widget(self.f_l60)
        else:
            self.nkt_box_wrap.size_hint_x, self.nkt_box_wrap.opacity, self.nkt_box_wrap.disabled = 1, 1, False
            self.btn_gno.size_hint_x = 1
            
            self.f_h_v = self.create_small_field("Глубина подвески ГНО / Воронки (м)", "2480")
            self.gno_fields_box.add_widget(self.f_h_v)

        # ФИКС КРАША ПРИ СОХРАНЕНИИ: Сюда передаются ТОЛЬКО валидные свойства MDCard
        self.calc_anchor.clear_widgets()

        if "ЭЦН" not in self.selected_gno.upper() and self.v_pipe is None:
            # Карточка БЕЗ аргумента text
            attention_table = MDCard(
                orientation="vertical",
                padding=[dp(15), dp(12)],
                size_hint=(1, 1),
                radius=[dp(10)],
                md_bg_color=(0.85, 0.4, 0.0, 1),
                ripple_behavior=False
            )
            attention_table.size_hint_max_x = dp(2000)
            
            # Текст строго внутри MDLabel
            attention_text = MDLabel(
                text="[b]ВНИМАНИЕ![/b] СНАЧАЛА ВЫБЕРИТЕ ТИП НКТ ИЗ БАЗЫ EXCEL!",
                halign="center",
                valign="middle",
                markup=True,
                theme_text_color="Custom",
                text_color=(1, 1, 1, 1),
                font_style="Button"
            )
            attention_table.add_widget(attention_text)
            self.calc_anchor.add_widget(attention_table)
        else:
            self.btn_calc = MDFillRoundFlatIconButton(
                text="ВЫПОЛНИТЬ РАСЧЕТ ТЕХНОЛОГИИ", 
                icon="calculator", 
                size_hint=(1, 1), 
                md_bg_color=(0.1, 0.6, 0.2, 1), 
                on_release=self.do_calc
            )
            self.btn_calc.size_hint_max_x = dp(2000)
            self.calc_anchor.add_widget(self.btn_calc)



    def do_calc(self, *args):
        if "ЭЦН" not in self.selected_gno.upper() and self.v_pipe is None: return
        try:
            def val(f): return float(f.text) if (hasattr(f, 'text') and f.text) else 0.0
            
            D, L_ckod, L_hvost, H_tvd = val(self.f_d_col), val(self.f_ckod), val(self.f_l_hvost), val(self.f_h_tvd)
            p_pl, p_ust, p_zatr, rho_c = val(self.f_p_pl), val(self.f_p_ust), val(self.f_p_zatr), val(self.f_rho)
            Q = self.q_coeffs[self.vtulka][self.gear] * self.rpm * 60 
            
            h_ref = H_tvd if H_tvd > 0 else L_ckod
            if h_ref == 0: return

            # Авторасчет пластового давления, если поле оставлено пустым
            if p_pl == 0:
                p_izb_calc = p_ust if p_ust > 0 else p_zatr
                p_pl = round(((h_ref * rho_c) / 10) + p_izb_calc, 1)
                calculated_p_pl_flag = True
            else:
                calculated_p_pl_flag = False

            # Выбор коэффициента запаса Кз по категориям Ростехнадзора/Сургутнефтегаза
            if "1 КАТЕГОРИЯ" in self.well_category:
                k_zapas = 1.10  
            elif "2 КАТЕГОРИЯ" in self.well_category:
                k_zapas = 1.07  
            else:
                k_zapas = 1.05  

            if self.kill_method == "STAND":
                rho = (p_pl * 10.197 / h_ref * k_zapas)
            else:
                p_izb = p_ust if p_ust > 0 else p_zatr
                rho = round(rho_c + (p_izb * 10.197 / h_ref) + 0.05, 2)
            
            v_col_m = ((D - 15)**2) * 0.0007854
            v_hvost_m = ((114 - 15)**2) * 0.0007854
            
            if "ЭЦН" in self.selected_gno.upper():
                l73, l60 = val(self.f_l73), val(self.f_l60)
                Hv = l73 + l60
                v_trub = ((3.02 * l73) + (1.99 * l60)) / 1000
            else:
                Hv = val(self.f_h_v)
                v_trub = (self.v_pipe * Hv) / 1000
            
            v_zatrub = (v_col_m * Hv / 1000) - (v_trub * 1.35)
            v_pod = (v_col_m * ((L_ckod - L_hvost) - Hv) / 1000) + (v_hvost_m * L_hvost / 1000) if L_hvost > 0 else (v_col_m * (L_ckod - Hv) / 1000)
            
            gno_low = self.selected_gno.lower()
            if "уст." in gno_low: v_total = v_trub
            elif "снят" in gno_low: v_total = v_zatrub + v_pod
            else: v_total = v_trub + v_zatrub + v_pod
            
            # Сургутский нормативный аварийный избыток +3.0 м3 поверх объема скважины
            v_total_with_zap = round(v_total + 3.0, 2)
            t_min = (v_total_with_zap * 1000) / Q if Q > 0 else 0

            # Подробные IWCF/Отраслевые регламенты действий
            if self.kill_method == "STAND":
                instr = "Стандартная промывка скважины прямой циркуляцией. Контролируйте уровни мерников ЦС. При разнице затруба и долива >0.5 м³ — СТОП СПО!"
            elif self.kill_method == "DRILLER":
                instr = (
                    f"[color=#FFCC00][b]ЭТАП 0: ЗАКРЫТИЕ СКВАЖИНЫ[/b][/color]\n"
                    f"• Герметизация устья превентором после обнаружения признаков ГНВП (рост уровня в мернике, перелив, изменение параметров раствора).\n\n"
                    f"[color=#33FF33][b]ЭТАП 1: ПЕРВЫЙ ЦИКЛ ЦИРКУЛЯЦИИ (ВЫМЫВ)[/b][/color]\n"
                    f"• Вымывание пачки флюида старым раствором [b]{rho_c:.2f} г/см³[/b] через дроссель.\n"
                    f"• Давление в бурильных трубах поддерживай постоянным за счет регулирования штуцера.\n"
                    f"• [b]Цель:[/b] Полностью вытеснить приток на поверхность. Пачка покажется через [b]{v_trub:.2f} м³[/b].\n\n"
                    f"[color=#33FF33][b]ЭТАП 2: ВТОРОЙ ЦИКЛ ЦИРКУЛЯЦИИ (ГЛУШЕНИЕ)[/b][/color]\n"
                    f"• Закачивание утяжеленного раствора глушения плотностью [color=#FF5555][b]{rho:.2f} г/см³[/b][/color].\n"
                    f"• Давление в бурильных трубах снижай по мере продвижения утяжеленного раствора до забоя.\n"
                    f"• [b]Цель:[/b] Уравновесить пластовое давление и полностью заглушить скважину.\n\n"
                    f"[color=#888888]----------------------------------------\n"
                    f"[b]ПЛЮСЫ:[/b] Можно начинать немедленно (не ждать утяжеления).\n"
                    f"[b]МИНУСЫ:[/b] Требует минимум двух циклов. Создает более высокое давление на устье.[/color]"
                )
            elif self.kill_method == "WAIT":
                instr = (
                    f"[color=#33FF33][b]МЕТОД ОЖИДАНИЯ И УТЯЖЕЛЕНИЯ:[/b][/color]\n"
                    f"• Суть: Герметизация устья, расчет и приготовление утяжеленного раствора, вымывание флюида за один цикл циркуляции.\n\n"
                    f"[b]ЭТАПЫ ПРОВЕДЕНИЯ:[/b]\n"
                    f"1. [b]Герметизация и остановка:[/b] При фиксации ГНВП устье мгновенно закрывается превенторами.\n"
                    f"2. [b]Ожидание:[/b] Скважина закрыта, пока на буровой готовится раствор плотности [b]{rho:.2f} г/см³[/b] в количестве [b]{v_total_with_zap:.2f} м³[/b].\n"
                    f"3. [b]Закачка и вымывание:[/b] Тяжелый раствор закачивается в скважину, флюид вымывается через манифольд при постоянном контроле устьевого давления.\n\n"
                    f"[color=#888888][b]ПЛЮСЫ:[/b] Заглушение всего за один цикл. Снижает устьевое давление до минимума.\n"
                    f"[b]МИНУСЫ:[/b] Скважина долго простаивает под давлением во время утяжеления.[/color]"
                )
            else:
                instr = f"ЗАДАВКА (BULLHEADING):\nГлушение давлением без циркуляции. Качай Ро {rho:.2f} напрямую в затрубное пространство. Внимание: контроль давления гидроразрыва пласта!"

            res_scr = self.manager.get_screen('kill_res')
            res_scr.selected_gno = self.selected_gno
            p_pl_lbl = f"{p_pl} атм [color=#FFCC00](Расчет по устью)[/color]" if calculated_p_pl_flag else f"{p_pl} атм"

            res_scr.result_text.text = (
                f"[size=22sp][color=#33FF33]{self.selected_gno.upper()}[/color][/size]\n"
                f"[size=15sp][color=#FF5555]Уровень опасности: {self.well_category}[/color][/size]\n\n"
                f"• Плотность жидкости глушения (с коэффициентом): [color=#FF5555][b]{rho:.2f} г/см³[/b][/color]\n"
                f"• Пластовое давление: [b]{p_pl_lbl}[/b]\n"
                f"• Объем с запасом регламента (+3м³): [b]{v_total_with_zap:.2f} м³[/b]\n"
                f"• Время закачки агрегатом: [color=#FFFF33][b]{int(t_min)} мин[/b][/color]\n"
                f"• Инструмент НКТ: [b]{v_trub:.2f} м³[/b] | Кольцевое: [b]{v_zatrub:.2f} м³[/b]"
            )
            res_scr.instruction_text.text = instr
            self.manager.current = 'kill_res'
        except Exception as e: print(f"Error Monolith Calc: {e}")
    # --- МЕНЮ ВЫПАДАЮЩИХ СПИСКОВ С КРИТЕРИЯМИ КАТЕГОРИЙ СУРГУТНЕФТЕГАЗА ---
    def menu_category(self, b):
        items = [
            {"viewclass": "OneLineListItem", "text": "1 КАТЕГОРИЯ: Газовые; ГФ >= 200 м³/т; АВПД; H2S выше нормы", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_category_val("Скважина 1 категории!")},
            {"viewclass": "OneLineListItem", "text": "2 КАТЕГОРИЯ: Нефтяные с ГФ < 200 м³/т (Рпл > Ргидр до 15%); H2S в ПДК", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_category_val("Скважина 2 категории!")},
            {"viewclass": "OneLineListItem", "text": "3 КАТЕГОРИЯ: Рпл <= Ргидр (норма/пониж.); H2S полностью отсутствует", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_category_val("Скважина 3 категории!")}
        ]
        self.m_cat = MDDropdownMenu(caller=b, items=items, width_mult=6.5, md_bg_color=(0.1, 0.1, 0.15, 1))
        self.m_cat.open()


    def set_category_val(self, val):
        self.well_category = val
        if hasattr(self, 'm_cat') and self.m_cat: self.m_cat.dismiss()
        self.refresh_ui_elements()

    def menu_kill_methods(self, b):
        items = [
            {"viewclass": "OneLineListItem", "text": "ШТАТНЫЙ РЕЖИМ ГЛУШЕНИЯ", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_kill_method_from_menu("STAND")},
            {"viewclass": "OneLineListItem", "text": "МЕТОД БУРИЛЬЩИКА (2 ЦИКЛА)", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_kill_method_from_menu("DRILLER")},
            {"viewclass": "OneLineListItem", "text": "ОЖИДАНИЕ И УТЯЖЕЛЕНИЕ (1 ЦИКЛ)", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_kill_method_from_menu("WAIT")},
            {"viewclass": "OneLineListItem", "text": "ЗАДАВКА В ПЛАСТ (BULLHEADING)", "theme_text_color": "Custom", "text_color": (1,1,1,1), "on_release": lambda: self.set_kill_method_from_menu("BULL")}
        ]
        self.m_kill = MDDropdownMenu(caller=b, items=items, width_mult=5, md_bg_color=(0.1, 0.1, 0.15, 1))
        self.m_kill.open()

    def set_kill_method_from_menu(self, mode):
        self.kill_method = mode
        if hasattr(self, 'm_kill') and self.m_kill: self.m_kill.dismiss()
        self.refresh_ui_elements()

    def menu_gno(self, b):
        items = [{"viewclass":"OneLineListItem","text":i,"theme_text_color":"Custom","text_color":(1,1,1,1),"on_release":lambda x=i:self.set_gno(x)} for i in ["Воронка", "ЭЦН", "Пакер (уст.)", "Пакер (снят)"]]
        self.m_gno = MDDropdownMenu(caller=b, items=items, width_mult=4, md_bg_color=(0.1, 0.1, 0.15, 1)); self.m_gno.open()
        
    def set_gno(self, v): 
        self.selected_gno = v
        if hasattr(self, 'm_gno') and self.m_gno: self.m_gno.dismiss()
        self.refresh_ui_elements()

    def open_nkt_menu(self, b):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True); ws = wb.active
            head = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = [dict(zip(head, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r and r]
            items = [{"viewclass": "OneLineListItem", "text": f"НКТ {r['Тип_НКТ']}", "theme_text_color":"Custom","text_color":(1,1,1,1), "on_release": lambda x=r: self.set_nkt(x)} for r in rows]
            self.m_nkt = MDDropdownMenu(caller=b, items=items, width_mult=4, md_bg_color=(0.1, 0.1, 0.15, 1)); self.m_nkt.open()
        except Exception as e: print(f"Excel Err: {e}")

    def set_nkt(self, r): 
        try:
            d_vnutr = float(r.get('Внутр_d_мм', 62) or 62)
            self.v_pipe = (d_vnutr ** 2) * 0.0007854
            self.tech_nkt_name = str(r.get('Тип_НКТ', ''))
            if hasattr(self, 'm_nkt') and self.m_nkt: self.m_nkt.dismiss()
            self.refresh_ui_elements()
        except Exception as e: print(f"Set Nkt Err: {e}")

    def menu_v(self, b): self.m_v = MDDropdownMenu(caller=b, items=[{"viewclass":"OneLineListItem","text":i,"theme_text_color":"Custom","text_color":(1,1,1,1),"on_release":lambda x=i:self.set_v(x)} for i in ["100", "115", "127"]], width_mult=2, md_bg_color=(0.1, 0.1, 0.15, 1)); self.m_v.open()
    def set_v(self, v): self.vtulka = v; self.btn_v.text = f"ВТ:{v}"; self.m_v.dismiss()
    
    def menu_g(self, b): 
        items = [{"viewclass":"OneLineListItem","text":f"СК:{i}","theme_text_color":"Custom","text_color":(1,1,1,1),"on_release":lambda x=i:self.set_g(x)} for i in ["2","3","4","5"]]
        self.m_g = MDDropdownMenu(caller=b, items=items, width_mult=2, md_bg_color=(0.1, 0.1, 0.15, 1))
        self.m_g.open()

    def set_g(self, v): 
        self.gear = v
        self.btn_g.text = f"СКОР:{v}"
        if hasattr(self, 'm_g') and self.m_g: 
            self.m_g.dismiss()

    def menu_r(self, b): 
        items = [{"viewclass":"OneLineListItem","text":str(i),"theme_text_color":"Custom","text_color":(1,1,1,1),"on_release":lambda x=i:self.set_r(x)} for i in range(600, 2100, 100)]
        self.m_r = MDDropdownMenu(caller=b, items=items, width_mult=2, md_bg_color=(0.1, 0.1, 0.15, 1))
        self.m_r.open()

    def set_r(self, v): 
        self.rpm = v
        self.btn_r.text = f"ОБ:{v}"
        if hasattr(self, 'm_r') and self.m_r: 
            self.m_r.dismiss()

    def go_back(self): 
        self.manager.current = 'menu'

class SpoScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.v_metalla = 0
        self.v_vnutr = 0
        self.siphon_mode = False
        
        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))

        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(50),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        btn_back = MDIconButton(
            icon="arrow-left", pos_hint={'center_y': .5},
            theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: self.go_back()
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(
            text="КОНТРОЛЬ ДОЛИВА", theme_text_color="Custom",
            text_color=(1, 1, 1, 1), font_style="H6"
        ))
        main_layout.add_widget(self.toolbar)

        # 2. СКРОЛЛ
        scroll = ScrollView(size_hint=(1, 1))
        self.container = MDBoxLayout(
            orientation='vertical', spacing=dp(15), 
            padding=[dp(20), dp(20), dp(20), dp(20)], 
            size_hint_y=None, adaptive_height=True
        )

        # ФУНКЦИЯ ДЛЯ СОЗДАНИЯ ШИРОКИХ КНОПОК
        def create_wide_btn(text, color, func):
            anchor = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(52))
            btn = MDFillRoundFlatButton(
                text=text, size_hint=(1, 1), md_bg_color=color, on_release=func
            )
            # Снимаем ограничение ширины
            btn.size_hint_max_x = dp(2000)
            anchor.add_widget(btn)
            return anchor, btn

        # КНОПКА НКТ
        anchor_nkt, self.nkt_btn = create_wide_btn("ВЫБРАТЬ ТИП НКТ", (0.1, 0.3, 0.6, 1), self.open_nkt_menu)

        # КНОПКА СИФОНА
        anchor_siph, self.siphon_btn = create_wide_btn("БЕЗ СИФОНА (ТОЛЬКО МЕТАЛЛ)", (0.1, 0.6, 0.2, 1), self.toggle_siphon)

        # ПОЛЯ ВВОДА
        self.count = MDTextField(hint_text="Кол-во трубок / свечей (шт)", mode="fill", input_filter="float", size_hint_x=1)
        self.len_one = MDTextField(hint_text="Длина одной ед. (м)", text="10.1", mode="fill", input_filter="float", size_hint_x=1)

        # КНОПКА РАСЧЕТА
        anchor_calc = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(55))
        self.calc_btn = MDFillRoundFlatIconButton(
            text="РАССЧИТАТЬ ОБЪЕМ", icon="calculator", size_hint=(1, 1),
            md_bg_color=(0.1, 0.4, 0.8, 1), on_release=self.calculate_doliv
        )
        self.calc_btn.size_hint_max_x = dp(2000)
        anchor_calc.add_widget(self.calc_btn)

        # РЕЗУЛЬТАТ
        self.res_label = MDLabel(
            text="[color=#AAAAAA]Выберите НКТ[/color]", 
            theme_text_color="Custom", text_color=(1, 1, 1, 1),
            font_style="H5", size_hint_y=None, halign="center", markup=True
        )

        # КРУПНОЕ ПРЕДУПРЕЖДЕНИЕ ВНИЗУ
        self.warning_label = MDLabel(
            text="[color=#FFCC00][b]ВАЖНО:[/b][/color]\nПеред СПО убедись в наличии солевого раствора не менее [b]4,5 м³[/b]",
            theme_text_color="Custom", text_color=(1, 1, 1, 1),
            font_style="Subtitle1", halign="center", markup=True,
            size_hint_y=None, line_height=1.2
        )
        self.warning_label.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1] + dp(30)))

        # Сборка
        for w in [anchor_nkt, anchor_siph, self.count, self.len_one, anchor_calc, self.res_label, self.warning_label]:
            self.container.add_widget(w)

        scroll.add_widget(self.container)
        main_layout.add_widget(scroll)
        self.add_widget(main_layout)

    def toggle_siphon(self, btn):
        self.siphon_mode = not self.siphon_mode
        btn.text = "С СИФОНОМ (МЕТАЛЛ + ВНУТР.)" if self.siphon_mode else "БЕЗ СИФОНА (ТОЛЬКО МЕТАЛЛ)"
        btn.md_bg_color = (0.8, 0.1, 0.1, 1) if self.siphon_mode else (0.1, 0.6, 0.2, 1)
        if self.v_metalla > 0: self.calculate_doliv()

    def open_nkt_menu(self, button):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
            ws = wb.active
            head = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = [dict(zip(head, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
            items = [{"viewclass": "OneLineListItem", "text": f"НКТ {r['Тип_НКТ']}", "on_release": lambda x=r: self.set_item(x)} for r in rows]
            self.menu = MDDropdownMenu(caller=button, items=items, width_mult=4)
            self.menu.open()
        except: self.res_label.text = "[color=#FF5555]Ошибка базы[/color]"

    def set_item(self, row):
        self.nkt_btn.text = f"НКТ {row['Тип_НКТ']}"
        self.v_metalla = float(row.get('V_metalla_m3', 0))
        self.v_vnutr = float(row.get('V_vnutr_m3', 0))
        self.menu.dismiss()
        self.calculate_doliv()

    def calculate_doliv(self, *args):
        try:
            if self.v_metalla == 0: return
            cnt = float(self.count.text or 0)
            l_one = float(self.len_one.text or 0)
            v_factor = (self.v_metalla + self.v_vnutr) if self.siphon_mode else self.v_metalla
            vol = (cnt * l_one) * v_factor * 1000
            self.res_label.text = f"[size=24sp][color=#33FF33]ДОЛИВ: {vol:.1f} л[/color][/size]"
        except: self.res_label.text = "Ошибка"

    def go_back(self):
        self.manager.current = 'menu'

class VzdScreen(BaseScreen):    # ВЗД и Режимы. Расчет оборотов винтового двигателя и справочник по долотам/фрезам
    def __init__(self, **kw):
        super().__init__(**kw)
        self.vtulka = "115"
        self.gear = "3"
        self.rpm_ca = 1000
        
        # Коэффициенты подачи Q для ЦА-320 (л/об)
        self.q_coeffs = {
            "100": {"2": 0.171, "3": 0.342, "4": 0.513, "5": 0.777},
            "115": {"2": 0.227, "3": 0.453, "4": 0.680, "5": 1.020},
            "127": {"2": 0.277, "3": 0.555, "4": 0.832, "5": 1.247}
        }
        
        # РАСШИРЕННАЯ БАЗА ВЗД (из техпаспортов)
        self.vzd_data = {
            "Д-76": {"flow": (4, 8), "ratio": 1.25, "torque": 850, "wob": 4.5, "press": 40},
            "Д-85": {"flow": (6, 12), "ratio": 0.88, "torque": 1300, "wob": 6.0, "press": 50},
            "Д-106": {"flow": (8, 16), "ratio": 0.65, "torque": 2200, "wob": 8.0, "press": 65},
            "Д-121": {"flow": (10, 22), "ratio": 0.52, "torque": 2900, "wob": 10.0, "press": 70}
        }
        self.selected_vzd = "Д-85"

        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))

        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(50),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        from kivymd.uix.button import MDIconButton
        btn_back = MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: self.go_back()
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(text="ВЗД И ИНСТРУМЕНТ", theme_text_color="Custom", text_color=(1, 1, 1, 1), font_style="H6"))
        main_layout.add_widget(self.toolbar)

        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(12), padding=dp(15))

        # 2. ВЫБОР ДВИГАТЕЛЯ
        self.btn_vzd = MDFillRoundFlatButton(
            text=f"ДВИГАТЕЛЬ: {self.selected_vzd}", 
            size_hint=(1, None), height=dp(48), md_bg_color=(0.1, 0.3, 0.6, 1),
            on_release=self.menu_vzd
        )
        self.container.add_widget(self.btn_vzd)

        # 3. РЕЖИМ НАСОСА
        self.container.add_widget(MDLabel(text="РЕЖИМ АГРЕГАТА ЦА-320:", font_style="Button", theme_text_color="Hint", halign="center"))
        pump_grid = MDGridLayout(cols=3, spacing=dp(5), size_hint_y=None, height=dp(45))
        
        self.btn_v = MDFillRoundFlatButton(text=f"ВТ:{self.vtulka}", size_hint=(1,1), md_bg_color=(0.1, 0.4, 0.8, 1), on_release=self.menu_v)
        self.btn_g = MDFillRoundFlatButton(text=f"СКОР:{self.gear}", size_hint=(1,1), md_bg_color=(0.1, 0.4, 0.8, 1), on_release=self.menu_g)
        self.btn_r = MDFillRoundFlatButton(text=f"ОБ:{self.rpm_ca}", size_hint=(1,1), md_bg_color=(0.1, 0.4, 0.8, 1), on_release=self.menu_r)
        
        for b in [self.btn_v, self.btn_g, self.btn_r]: pump_grid.add_widget(b)
        self.container.add_widget(pump_grid)

        # 4. ПЕРЕКЛЮЧАТЕЛЬ ИНФОРМАЦИИ
        type_grid = MDGridLayout(cols=3, spacing=dp(5), size_hint_y=None, height=dp(45))
        self.b1 = MDFillRoundFlatButton(text="РЕЖИМ", size_hint=(1,1), on_release=lambda x: self.switch_info("ВЗД"))
        self.b2 = MDFillRoundFlatButton(text="ДОЛОТА", size_hint=(1,1), on_release=lambda x: self.switch_info("ДОЛОТО"))
        self.b3 = MDFillRoundFlatButton(text="ФРЕЗЫ", size_hint=(1,1), on_release=lambda x: self.switch_info("ФРЕЗ"))
        for b in [self.b1, self.b2, self.b3]: type_grid.add_widget(b)
        self.container.add_widget(type_grid)

        # 5. ОКНО ВЫВОДА (С поддержкой разметки)
        self.res_label = MDLabel(
            text="Загрузка данных...", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            font_style="Body1", size_hint_y=None, markup=True, line_height=1.2
        )
        self.res_label.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1]))
        self.container.add_widget(self.res_label)

        scroll.add_widget(self.container)
        main_layout.add_widget(scroll)
        self.add_widget(main_layout)
        self.switch_info("ВЗД")

    # --- ЛОГИКА ---
    def switch_info(self, mode):
        active, inactive = (0.1, 0.4, 0.8, 1), (0.2, 0.2, 0.2, 1)
        self.b1.md_bg_color = active if mode == "ВЗД" else inactive
        self.b2.md_bg_color = active if mode == "ДОЛОТО" else inactive
        self.b3.md_bg_color = active if mode == "ФРЕЗ" else inactive
        if mode == "ВЗД": self.calc_vzd()
        elif mode == "ДОЛОТО": self.show_bits()
        else: self.show_mills()

    def calc_vzd(self):
        # Расчет подачи в л/с
        Q = (self.q_coeffs[self.vtulka][self.gear] * self.rpm_ca) / 60
        v = self.vzd_data[self.selected_vzd]
        rpm_vzd = int(Q * v["ratio"] * 60)
        
        # Цветовая индикация расхода
        status_color = "#33FF33" if v["flow"][0] <= Q <= v["flow"][1] else "#FF5555"
        status_text = "В НОРМЕ" if status_color == "#33FF33" else "ВНЕ ДИАПАЗОНА!"

        self.res_label.text = (
            f"[size=20sp][color=#33FF33]ТЕХ.КАРТА {self.selected_vzd}:[/color][/size]\n\n"
            f"• Подача ЦА: [b]{Q:.2f} л/с[/b] ({status_text})\n"
            f"• Обороты: [color={status_color}][b]{rpm_vzd} об/мин[/b][/color]\n"
            f"• Допуст. расход: [b]{v['flow'][0]}-{v['flow'][1]} л/с[/b]\n"
            f"• Макс. нагрузка (WOB): [b]{v['wob']} тонн[/b]\n"
            f"• Дифф. давление ($\Delta P$): [b]{v['press']/10:.1f} МПа ({v['press']} атм)[/b]\n"
            f"--------------------------------\n"
            f"[color=#FFCC00][b]ПРОВЕРКА НА УСТЬЕ:[/b][/color]\n"
            f"1. Подать миним. расход [b]{v['flow'][0]} л/с[/b].\n"
            f"2. Мотор должен запуститься плавно, без рывков.\n"
            f"3. Давление ХХ должно быть стабильным.\n"
            f"--------------------------------\n"
            f"[color=#FFCC00][b]ПРИЗНАКИ ПОЛОМКИ:[/b][/color]\n"
            f"• Давление растет, а проходки нет — [b]ЗАКЛИНКА[/b].\n"
            f"• Давление упало, обороты выросли — [b]ПРОМЫВ СТАТОРА[/b].\n"
            f"• Стук в колонне — [b]ИЗНОС КАРДАНА[/b]."
        )

    def show_bits(self):
        self.res_label.text = (
            "[size=20sp][color=#33FF33]ИНСТРУКЦИЯ: ДОЛОТА[/color][/size]\n\n"
            "[b]PDC (Алмазные):[/b]\n"
            "• Применяются на высоких оборотах (>100 об/мин).\n"
            "• [color=#FF5555]Запрещено:[/color] работа по металлу и «прыжки».\n"
            "• Нагрузка: плавная, 1-3 тонны.\n\n"
            "[b]Шарошечные (Т, СТ):[/b]\n"
            "• Эффективны на малых оборотах с высокой нагрузкой.\n"
            "• Нагрузка: до [b]6-8 тонн[/b] (зависит от Ø).\n"
            "• Обязательна приработка подшипников (15 мин)."
        )

    def show_mills(self):
        self.res_label.text = (
            "[size=20sp][color=#33FF33]ИНСТРУКЦИЯ: ФРЕЗЫ[/color][/size]\n\n"
            "[b]Торцевые (ФТ):[/b]\n"
            "• Работа по «голове» аварийного инструмента.\n"
            "• Обязательна интенсивная промывка для выноса стружки.\n\n"
            "[b]Магнитные (ФМ):[/b]\n"
            "• Сбор мелкого металла (зубья, сухари).\n"
            "• Спускать без вращения ВЗД (на прямую промывку)."
        )

    # --- МЕНЮ ВЫБОРА ---
    def menu_vzd(self, b):
        items = [{"viewclass":"OneLineListItem","text":k,"on_release":lambda x=k: self.set_vzd(x)} for k in self.vzd_data.keys()]
        self.m_vzd = MDDropdownMenu(caller=b, items=items, width_mult=4); self.m_vzd.open()
    def set_vzd(self, v): self.selected_vzd = v; self.btn_vzd.text = f"ДВИГАТЕЛЬ: {v}"; self.m_vzd.dismiss(); self.calc_vzd()

    def menu_v(self, b):
        items = [{"viewclass":"OneLineListItem","text":i,"on_release":lambda x=i:self.set_v(x)} for i in ["100", "115", "127"]]
        self.m_v = MDDropdownMenu(caller=b, items=items, width_mult=2); self.m_v.open()
    def set_v(self, v): self.vtulka = v; self.btn_v.text = f"ВТ:{v}"; self.m_v.dismiss(); self.calc_vzd()

    def menu_g(self, b):
        items = [{"viewclass":"OneLineListItem","text":f"СК:{i}","on_release":lambda x=i:self.set_g(x)} for i in ["2","3","4","5"]]
        self.m_g = MDDropdownMenu(caller=b, items=items, width_mult=2); self.m_g.open()
    def set_g(self, v): self.gear = v; self.btn_g.text = f"СКОР:{v}"; self.m_g.dismiss(); self.calc_vzd()

    def menu_r(self, b):
        items = [{"viewclass":"OneLineListItem","text":str(i),"on_release":lambda x=i:self.set_r(x)} for i in range(600, 2100, 100)]
        self.m_r = MDDropdownMenu(caller=b, items=items, width_mult=2); self.m_r.open()
    def set_r(self, v): self.rpm_ca = v; self.btn_r.text = f"ОБ:{v}"; self.m_r.dismiss(); self.calc_vzd()

    def go_back(self): self.manager.current = 'menu'

class ThreadsScreen(BaseScreen):      # Справочник инструмента. Данные по НКТ, СБТ и пакерам (нагрузки, веса, объемы).
    def __init__(self, **kw):
        super().__init__(**kw)
        self.sel_type = "НКТ"
        self.sel_size = "73 x 5.5"
        self.sel_grade = "Д"
        
        # Группы прочности сталей (Предел текучести в МПа)
        self.steel_grades = {"Д": 379, "К": 492, "Е": 552, "Л": 655, "М": 758, "Р": 862}
        
        layout = MDBoxLayout(orientation='vertical')
        
        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(45),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        btn_back = MDIconButton(
            icon="arrow-left", pos_hint={'center_y': .5},
            theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: setattr(self.manager, 'current', 'menu')
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(text="СПРАВОЧНИК ИНСТРУМЕНТА", theme_text_color="Custom", 
                                       text_color=(1,1,1,1), font_style="H6"))
        layout.add_widget(self.toolbar)

        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(12), padding=dp(15))

        # 2. ТАБЫ (НКТ / СБТ / ПАКЕР)
        type_grid = MDGridLayout(cols=3, spacing=dp(5), size_hint_y=None, height=dp(45))
        
        self.btn_tab_nkt = MDFillRoundFlatButton(text="НКТ", size_hint=(1, 1), on_release=lambda x: self.switch_tab("НКТ"))
        self.btn_tab_sbt = MDFillRoundFlatButton(text="СБТ", size_hint=(1, 1), on_release=lambda x: self.switch_tab("СБТ"))
        self.btn_tab_pak = MDFillRoundFlatButton(text="ПАКЕР", size_hint=(1, 1), on_release=lambda x: self.switch_tab("ПАКЕР"))
        
        for b in [self.btn_tab_nkt, self.btn_tab_sbt, self.btn_tab_pak]: type_grid.add_widget(b)
        self.container.add_widget(type_grid)

        # 3. КНОПКИ ПАРАМЕТРОВ
        self.anchor_size = MDAnchorLayout(anchor_x='center', size_hint=(1, None), height=dp(50))
        self.btn_size = MDFillRoundFlatButton(text="РАЗМЕР", on_release=self.menu_sizes, size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1))
        self.anchor_size.add_widget(self.btn_size)

        self.anchor_steel = MDAnchorLayout(anchor_x='center', size_hint=(1, None), height=dp(50))
        self.btn_steel = MDFillRoundFlatButton(text="ГРУППА ПРОЧНОСТИ", on_release=self.menu_steel, size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1))
        self.anchor_steel.add_widget(self.btn_steel)
        
        self.container.add_widget(self.anchor_size)
        self.container.add_widget(self.anchor_steel)

        # 4. ВЫВОД ИНФОРМАЦИИ
        self.res_card = MDLabel(
            text="Загрузка...", theme_text_color="Custom", text_color=(1,1,1,1),
            font_style="Body1", size_hint_y=None, markup=True, line_height=1.3
        )
        self.res_card.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1]))
        self.container.add_widget(self.res_card)

        scroll.add_widget(self.container)
        layout.add_widget(scroll)
        self.add_widget(layout)
        self.switch_tab("НКТ")

    # --- ЛОГИКА ---
    def switch_tab(self, mode):
        self.sel_type = mode
        active, inactive = (0.1, 0.4, 0.8, 1), (0.2, 0.2, 0.2, 1)
        self.btn_tab_nkt.md_bg_color = active if mode == "НКТ" else inactive
        self.btn_tab_sbt.md_bg_color = active if mode == "СБТ" else inactive
        self.btn_tab_pak.md_bg_color = active if mode == "ПАКЕР" else inactive
        
        if mode == "ПАКЕР":
            self.sel_size = "ПМР-122"
            self.anchor_steel.opacity, self.anchor_steel.disabled = 0, True
        else:
            self.sel_size = "73 x 5.5" if mode == "НКТ" else "СБТ 73"
            self.anchor_steel.opacity, self.anchor_steel.disabled = 1, False
        self.update_info()

    def menu_sizes(self, b):
        if self.sel_type == "НКТ": 
            list_s = ["60 x 5.0", "73 x 5.5", "89 x 6.5", "НКТУ 89 x 7.3"]
        elif self.sel_type == "СБТ": 
            list_s = ["СБТ 60", "СБТ 73", "СБТ 89"]
        else: 
            list_s = ["ПМР-122", "ПМР-140", "ПРО-ЯДЖО-О", "2ПНМ-122"]
        items = [{"viewclass": "OneLineListItem", "text": i, "on_release": lambda x=i: self.set_size(x)} for i in list_s]
        self.m_size = MDDropdownMenu(caller=b, items=items, width_mult=4); self.m_size.open()

    def set_size(self, v):
        self.sel_size = v; self.m_size.dismiss(); self.update_info()

    def menu_steel(self, b):
        items = [{"viewclass": "OneLineListItem", "text": f"Сталь: {k}", "on_release": lambda x=k: self.set_grade(x)} for k in self.steel_grades.keys()]
        self.m_steel = MDDropdownMenu(caller=b, items=items, width_mult=3); self.m_steel.open()

    def set_grade(self, v):
        self.sel_grade = v; self.m_steel.dismiss(); self.update_info()

    def update_info(self):
        self.btn_size.text = f"РАЗМЕР: {self.sel_size}"
        self.btn_steel.text = f"СТАЛЬ: {self.sel_grade}"
        self.res_card.text = self.get_packer_info(self.sel_size) if self.sel_type == "ПАКЕР" else self.get_pipe_info()

    def get_pipe_info(self):
        # Расширенная база: {площадь_мм2, вес_кг, объем_л_м, резьба, момент_нм}
        data = {
            "60 x 5.0": {"area": 864, "weight": 7.1, "vol": 1.96, "thread": "60 ГОСТ 633", "torque": "2500-3500"},
            "73 x 5.5": {"area": 1166, "weight": 9.5, "vol": 3.02, "thread": "73 ГОСТ 633", "torque": "4500-5500"},
            "89 x 6.5": {"area": 1685, "weight": 13.5, "vol": 4.54, "thread": "89 ГОСТ 633", "torque": "7000-8000"},
            "НКТУ 89 x 7.3": {"area": 1870, "weight": 15.2, "vol": 4.35, "thread": "89-У (Усиленная)", "torque": "9000-11000"},
            "СБТ 60": {"area": 1100, "weight": 10.2, "vol": 1.45, "thread": "З-73 (NC26)", "torque": "5500-7500"},
            "СБТ 73": {"area": 1500, "weight": 14.5, "vol": 2.15, "thread": "З-86 (NC31)", "torque": "8000-10000"},
            "СБТ 89": {"area": 2100, "weight": 19.8, "vol": 3.42, "thread": "З-102 (NC38)", "torque": "12000-15000"}
        }
        d = data.get(self.sel_size, data["73 x 5.5"])
        sigma = self.steel_grades[self.sel_grade]
        force = (d["area"] * sigma) / 9810 * 0.8 # Тонны с запасом 0.8
        
        res = f"[size=22sp][color=#33FF33]{self.sel_size}[/color][/size]\n\n"
        res += f"• Разрывное (запас 0.8): [color=#FF5555][b]{force:.1f} тонн[/b][/color]\n"
        res += f"• Резьба: [b]{d['thread']}[/b]\n"
        res += f"• Момент затяжки: [b]{d['torque']} Н*м[/b]\n"
        res += f"• Вес метра: [b]{d['weight']} кг/м[/b] | Объем: [b]{d['vol']} л/м[/b]\n"
        res += f"• Группа прочности: [b]{self.sel_grade}[/b] ({sigma} МПа)\n"
        
        if "НКТУ" in self.sel_size:
            res += "--------------------------------\n[i]Применяется при ГРП и МГРП. Имеет увеличенную толщину стенки и усиленную муфту.[/i]"
        return res

    def get_packer_info(self, name):
        # Подробные инструкции по эксплуатации
        packers = {
            "ПМР-122": (
                "[b]Инструкция ПМР (Механический):[/b]\n"
                "1. Спустить до глубины. Выше цели на 1-2м.\n"
                "2. Приподнять на 0.5м, повернуть влево на 1/4 оборота.\n"
                "3. Разгрузить вес [b]8-10 тонн[/b]. Контроль по ИВЭ.\n"
                "• [color=#FFFF33]Снятие:[/color] Прямая натяжка вверх."
            ),
            "ПМР-140": (
                "[b]Инструкция ПМР-140:[/b]\n"
                "1. Спуск в колонну 168 мм.\n"
                "2. Активация осевым перемещением с поворотом.\n"
                "3. Рекомендуемая нагрузка [b]10-12 тонн[/b].\n"
                "• [color=#FFFF33]Внимание:[/color] Перед снятием выровнять давление."
            ),
            "ПРО-ЯДЖО-О": (
                "[b]Инструкция ПРО-ЯДЖО (Якорный):[/b]\n"
                "1. Спуск до интервала.\n"
                "2. Поворот вправо на 1/4 оборота.\n"
                "3. Натяжка вверх или разгрузка (зависит от исполнения).\n"
                "4. Удерживает перепад до [b]35-50 МПа[/b].\n"
                "• [color=#FFFF33]Снятие:[/color] Осевое перемещение в нейтраль."
            ),
            "2ПНМ-122": (
                "[b]Инструкция 2ПНМ (Гидравлический):[/b]\n"
                "1. Спустить на проектную глубину.\n"
                "2. Бросить шар (если предусмотрено) или создать Р.\n"
                "3. Поднять давление в трубах до [b]120-150 атм[/b].\n"
                "4. Выдержать 5-10 мин для фиксации плашек.\n"
                "• [color=#FFFF33]Снятие:[/color] Натяжка с разрывом срезных винтов."
            )
        }
        return f"[size=22sp][color=#33FF33]{name}[/color][/size]\n\n" + packers.get(name, "Инструкция дополняется...")

class FishingScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.area_cm2 = 0
        
        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))
        
        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(50),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        self.toolbar.add_widget(MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: self.go_back()
        ))
        self.toolbar.add_widget(MDLabel(text="ЛОВИЛЬНЫЙ ЦЕНТР", theme_text_color="Custom", 
                                       text_color=(1,1,1,1), font_style="H6"))
        main_layout.add_widget(self.toolbar)

        # 2. ОСНОВНОЙ СКРОЛЛ
        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(15), padding=dp(20))

        # ТАБЫ (РЕЖИМЫ)
        tab_grid = MDGridLayout(cols=2, spacing=dp(5), size_hint_y=None, height=dp(45))
        self.b1 = MDFillRoundFlatButton(text="РАСЧЕТ ПРИХВАТА", size_hint=(1,1), on_release=lambda x: self.switch_mode("CALC"))
        self.b2 = MDFillRoundFlatButton(text="ИНСТРУМЕНТАРИЙ", size_hint=(1,1), on_release=lambda x: self.switch_mode("TOOL"))
        tab_grid.add_widget(self.b1); tab_grid.add_widget(self.b2)
        self.container.add_widget(tab_grid)

        # КОНТЕНТ
        self.content_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(15))
        self.container.add_widget(self.content_box)

        scroll.add_widget(self.container); main_layout.add_widget(scroll); self.add_widget(main_layout)
        self.switch_mode("CALC")

    def switch_mode(self, mode):
        self.content_box.clear_widgets()
        active, inactive = (0.1, 0.4, 0.8, 1), (0.2, 0.2, 0.2, 1)
        self.b1.md_bg_color = active if mode == "CALC" else inactive
        self.b2.md_bg_color = active if mode == "TOOL" else inactive
        if mode == "CALC": self.draw_calc_ui()
        else: self.draw_tool_info()

    def draw_calc_ui(self):
        # Инструкция (Кратко и понятно)
        help_t = (
            "[color=#FFCC00][b]ЗАМЕР ВЫТЯЖКИ:[/b][/color]\n"
            "• Натяни до собственного веса, поставь метку.\n"
            "• Дай натяжку выше веса на [b]P[/b] (10-15 тн).\n"
            "• Замерь вытяжку трубы [b]dL[/b] в сантиметрах."
        )
        self.content_box.add_widget(MDLabel(text=help_t, markup=True, theme_text_color="Secondary", size_hint_y=None, adaptive_height=True))

        # Кнопка НКТ
        anchor_nkt = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(50))
        self.btn_nkt = MDFillRoundFlatButton(text="1. ВЫБРАТЬ НКТ", size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), on_release=self.open_nkt_menu)
        self.btn_nkt.size_hint_max_x = dp(2000); anchor_nkt.add_widget(self.btn_nkt)
        self.content_box.add_widget(anchor_nkt)

        self.f_force = MDTextField(hint_text="Натяжка P (тонн)", mode="fill", input_filter="float")
        self.f_stretch = MDTextField(hint_text="Вытяжка dL (сантиметры)", mode="fill", input_filter="float")
        self.content_box.add_widget(self.f_force); self.content_box.add_widget(self.f_stretch)

        # Кнопка расчета
        anchor_calc = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(55))
        btn_calc = MDFillRoundFlatIconButton(text="РАСЧИТАТЬ ГЛУБИНУ", icon="calculator", size_hint=(1, 1), md_bg_color=(0.1, 0.6, 0.2, 1), on_release=self.calc_stuck)
        btn_calc.size_hint_max_x = dp(2000); anchor_calc.add_widget(btn_calc)
        self.content_box.add_widget(anchor_calc)

        # Табло результата (Карточка)
        self.res_card = MDCard(orientation='vertical', padding=dp(15), size_hint_y=None, adaptive_height=True, md_bg_color=(0,0,0,1))
        self.res_label = MDLabel(text="Верхняя граница прихвата: -- м", halign="center", theme_text_color="Custom", text_color=(0.2, 1, 0.2, 1), bold=True)
        self.res_card.add_widget(self.res_label)
        self.content_box.add_widget(self.res_card)

    def draw_tool_info(self):
        info = (
            "[size=18sp][color=#33FF33]БАЗА ЛОВИЛЬНОГО ИНСТРУМЕНТА:[/color][/size]\n\n"
            "[b]1. ОВЕРШОТЫ (ТИП ОС):[/b]\n"
            "• Назначение: Наружный захват за муфту/тело.\n"
            "• [b]Захват:[/b] Спиральные или цанговые граппы.\n"
            "• [b]Метод:[/b] Накрыть «голову», посадка 1-2т, поворот вправо 1/4, натяжка.\n\n"
            "[b]2. МЕТЧИКИ (МЭС / МЭУ):[/b]\n"
            "• Назначение: Внутренний захват за резьбу или тело.\n"
            "• [b]МЭС:[/b] Специальная резьба для врезки.\n"
            "• [b]Метод:[/b] Заворот вправо 5-8 оборотов. ОСТОРОЖНО: Инструмент хрупкий!\n\n"
            "[b]3. КОЛОКОЛА (К / КГ):[/b]\n"
            "• Назначение: Нарезка наружной резьбы на инструменте.\n"
            "• [b]Метод:[/b] Вращение 15-20 об/мин с осевой нагрузкой 1.5-2.0 тн.\n\n"
            "[b]4. ПЕЧАТИ (П-140 / П-118):[/b]\n"
            "• Определение формы «головы» и её положения в колонне.\n"
            "• [b]Важно:[/b] Посадка [b]не более 2-3 тонн[/b], чтобы не смять корпус."
        )
        self.content_box.add_widget(MDLabel(text=info, markup=True, theme_text_color="Custom", text_color=(1,1,1,1), size_hint_y=None, adaptive_height=True))

    def calc_stuck(self, *args):
        try:
            if self.area_cm2 == 0:
                self.res_label.text = "Сначала выберите НКТ!"; return
            P = float(self.f_force.text or 0) * 1000 # тн -> кг
            L = float(self.f_stretch.text or 0)      # см
            if P <= 0: return
            # Формула: H = (E * F * dL) / P. E=2.1e6
            depth = (2100000 * self.area_cm2 * L) / P / 100 # см -> м
            self.res_label.text = f"Верх прихвата: ~{int(depth)} метров"
        except: self.res_label.text = "Ошибка в данных!"

    def open_nkt_menu(self, b):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True); ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = [dict(zip(headers, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]
            items = [{"viewclass": "OneLineListItem", "text": f"НКТ {r['Тип_НКТ']}", "on_release": lambda x=r: self.set_nkt(x)} for r in rows]
            self.menu = MDDropdownMenu(caller=b, items=items, width_mult=4); self.menu.open()
        except: self.res_label.text = "Ошибка базы"

    def set_nkt(self, r):
        # S сечения = V_metalla * 10000 (перевод м2 в см2)
        self.area_cm2 = float(r.get('V_metalla_m3', 0)) * 10000
        self.btn_nkt.text = f"ИНСТРУМЕНТ: {r['Тип_НКТ']}"; self.menu.dismiss()

    def go_back(self): self.manager.current = 'menu'

class SettingsScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.current_version = "1.0.2"
        self.update_url = "https://github.com"
        
        layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))
        
        # 1. ШАПКА ЭКРАНА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(45),
            md_bg_color=(0.2, 0.25, 0.3, 1), padding=[dp(15), 0, dp(10), 0]
        )
        self.toolbar.add_widget(MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            pos_hint={'center_y': .5},
            on_release=lambda x: setattr(self.manager, 'current', 'menu')
        ))
        self.toolbar.add_widget(MDLabel(text="НАСТРОЙКИ ПРИЛОЖЕНИЯ", theme_text_color="Custom", text_color=(1, 1, 1, 1), font_style="H6"))
        layout.add_widget(self.toolbar)
        
        # 2. СКРОЛЛ КОНТЕНТА
        scroll = ScrollView()
        container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(12), padding=dp(20))
        
        container.add_widget(MDLabel(text="КОЭФФИЦИЕНТЫ ЗАПАСА Ростехнадзора (Кз):", font_style="Caption", theme_text_color="Hint"))
        
        # Поля ввода для Кз
        self.f_k_z1 = self.create_setting_field("Кз для 1 КАТЕГОРИИ (Газ/АВПД)", "1.10")
        self.f_k_z2 = self.create_setting_field("Кз для 2 КАТЕГОРИИ (Нефтяные)", "1.07")
        self.f_k_z3 = self.create_setting_field("Кз для 3 КАТЕГОРИИ (Обычные)", "1.05")
        
        container.add_widget(self.f_k_z1)
        container.add_widget(self.f_k_z2)
        container.add_widget(self.f_k_z3)
        
        container.add_widget(MDLabel(text="ТЕХНОЛОГИЧЕСКИЕ РЕГЛАМЕНТЫ КРС:", font_style="Caption", theme_text_color="Hint"))
        
        # Поля ввода для нормативных запасов
        self.f_v_zapas = self.create_setting_field("Аварийный запас объема на мерник (м³)", "3.0")
        self.f_v_crit_spo = self.create_setting_field("Допуск на перелив при СПО (м³)", "0.5")
        
        container.add_widget(self.f_v_zapas)
        container.add_widget(self.f_v_crit_spo)
        
        # 3. СТАТУС И ОБНОВЛЕНИЕ БАЗ ДАННЫХ EXCEL С ТВОЕГО GITHUB
        container.add_widget(MDLabel(text="УПРАВЛЕНИЕ ИНЖЕНЕРНЫМИ БАЗАМИ ДАННЫХ:", font_style="Caption", theme_text_color="Hint"))
        self.db_status_label = MDLabel(text="Базы данных: опрос...", theme_text_color="Secondary", font_style="Body2", size_hint_y=None, height=dp(25))
        container.add_widget(self.db_status_label)
        
        anchor_db_update = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(48))
        self.btn_db_update = MDFillRoundFlatIconButton(
            text="ОБНОВИТЬ ИНЖЕНЕРНЫЕ БАЗЫ С ГИТХАБ", icon="database-refresh",
            size_hint=(1, 1), md_bg_color=(0.1, 0.4, 0.8, 1),
            on_release=self.download_fresh_databases
        )
        self.btn_db_update.size_hint_max_x = dp(2000)
        anchor_db_update.add_widget(self.btn_db_update)
        container.add_widget(anchor_db_update)
        
        # 4. ОБНОВЛЕНИЕ СИСТЕМЫ (ПОИСК НОВЫХ APK)
        container.add_widget(MDLabel(text=f"ВЕРСИЯ ПРИЛОЖЕНИЯ: v{self.current_version}", font_style="Caption", theme_text_color="Hint"))
        
        anchor_update = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(48))
        self.btn_update_apk = MDFillRoundFlatIconButton(
            text="ПРОВЕРИТЬ ОБНОВЛЕНИЯ", icon="cloud-download-outline",
            size_hint=(1, 1), md_bg_color=(0.1, 0.4, 0.8, 1),
            on_release=self.check_for_updates
        )
        self.btn_update_apk.size_hint_max_x = dp(2000)
        anchor_update.add_widget(self.btn_update_apk)
        container.add_widget(anchor_update)
        
        container.add_widget(MDBoxLayout(size_hint_y=None, height=dp(10)))

        # 5. КНОПКА СОХРАНЕНИЯ (СТЕНА)
        anchor_save = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(55))
        self.btn_save_config = MDFillRoundFlatIconButton(
            text="СОХРАНИТЬ КОНФИГУРАЦИЮ", icon="content-save-outline",
            size_hint=(1, 1), md_bg_color=(0.1, 0.6, 0.2, 1),
            on_release=self.save_settings
        )
        self.btn_save_config.size_hint_max_x = dp(2000)
        anchor_save.add_widget(self.btn_save_config)
        container.add_widget(anchor_save)
        
        scroll.add_widget(container)
        layout.add_widget(scroll)
        self.add_widget(layout)

    def create_setting_field(self, hint, def_val):
        f = MDTextField(hint_text=hint, text=def_val, mode="fill", size_hint_y=None, height=dp(34), input_filter='float')
        f.bind(focus=lambda ins, val: Clock.schedule_once(lambda dt: ins.select_all(), 0.1) if val else None)
        return f

    def on_enter(self):
        store = JsonStore(os.path.join(BASE_DIR, 'settings.json'))
        if store.exists('engineering_specs'):
            data = store.get('engineering_specs')
            self.f_k_z1.text = str(data.get('kz1', '1.10'))
            self.f_k_z2.text = str(data.get('kz2', '1.07'))
            self.f_k_z3.text = str(data.get('kz3', '1.05'))
            self.f_v_zapas.text = str(data.get('v_zapas', '3.0'))
            self.f_v_crit_spo.text = str(data.get('v_crit_spo', '0.5'))
            
        # Опрашиваем физическое наличие всех трех файлов Excel в памяти устройства
        nkt_ok = os.path.exists(EXCEL_PATH)
        vzd_ok = os.path.exists(VZD_PATH)
        ca_ok = os.path.exists(os.path.join(BASE_DIR, 'data', 'ca320_stats.xlsx'))
        
        status_txt = (
            f"НКТ: {'[color=#33FF33]ОК[/color]' if nkt_ok else '[color=#FF3333]НЕТ[/color]'} | "
            f"ВЗД: {'[color=#33FF33]ОК[/color]' if vzd_ok else '[color=#FF3333]НЕТ[/color]'} | "
            f"ЦА-320: {'[color=#33FF33]ОК[/color]' if ca_ok else '[color=#FF3333]НЕТ[/color]'}"
        )
        self.db_status_label.text = status_txt
        self.db_status_label.markup = True

    def download_fresh_databases(self, *args):
        # Фоновое скачивание, чтобы интерфейс приложения не зависал намертво
        self.btn_db_update.text = "СКАЧИВАНИЕ ТАБЛИЦ С ГИТХАБ..."
        self.btn_db_update.disabled = True
        import threading
        threading.Thread(target=self._run_download).start()

    def _run_download(self):
        import urllib.request
        try:
            # 1. Загрузка базы НКТ
            urllib.request.urlretrieve(
                "https://githubusercontent.com", 
                EXCEL_PATH
            )
            # 2. Загрузка каталога ВЗД
            urllib.request.urlretrieve(
                "https://githubusercontent.com", 
                VZD_PATH
            )
            # 3. Загрузка статистики ЦА-320
            ca_path = os.path.join(BASE_DIR, 'data', 'ca320_stats.xlsx')
            urllib.request.urlretrieve(
                "https://githubusercontent.com", 
                ca_path
            )
            Clock.schedule_once(lambda dt: self._download_success(), 0)
        except Exception as e:
            # ФИКС NAMEERROR: жестко переводим ошибку в текст до того, как Python сотрёт переменную 'e'
            err_msg = str(e)
            Clock.schedule_once(lambda dt: self._download_error(err_msg), 0)


    def _download_success(self):
        self.btn_db_update.text = "БАЗЫ ДАННЫХ УСПЕШНО ОБНОВЛЕНЫ!"
        self.btn_db_update.md_bg_color = (0.1, 0.6, 0.2, 1) # Зеленый цвет успеха
        self.btn_db_update.disabled = False
        self.on_enter() # Обновит статус-маркеры на экране
        
        def restore_db_btn(dt):
            self.btn_db_update.text = "ОБНОВИТЬ ИНЖЕНЕРНЫЕ БАЗЫ С ГИТХАБ"
            self.btn_db_update.md_bg_color = (0.1, 0.4, 0.8, 1)
        Clock.schedule_once(restore_db_btn, 3.0)

    def _download_error(self, err_msg):
        self.btn_db_update.text = "ОШИБКА ЗАГРУЗКИ (ПРОВЕРЬ СВЯЗЬ)"
        self.btn_db_update.md_bg_color = (0.85, 0.15, 0.15, 1) # Красный цвет ошибки
        self.btn_db_update.disabled = False
        
        def restore_db_btn(dt):
            self.btn_db_update.text = "ОБНОВИТЬ ИНЖЕНЕРНЫЕ БАЗЫ С ГИТХАБ"
            self.btn_db_update.md_bg_color = (0.1, 0.4, 0.8, 1)
        Clock.schedule_once(restore_db_btn, 4.0)
        print(f"Update error: {err_msg}")

    def check_for_updates(self, *args):
        # Интерактивный отклик на кнопке обновлений
        self.btn_update_apk.text = "ПЕРЕХОД НА СТРАНИЦУ ЗАГРУЗКИ..."
        self.btn_update_apk.md_bg_color = (0.1, 0.5, 0.8, 1) # Синий цвет перехода
        
        def restore_update_btn(dt):
            self.btn_update_apk.text = "ПРОВЕРИТЬ ОБНОВЛЕНИЯ"
            self.btn_update_apk.md_bg_color = (0.1, 0.4, 0.8, 1)
        Clock.schedule_once(restore_update_btn, 3.0)

        try:
            webbrowser.open(self.update_url)
        except Exception as e:
            print(f"Browser error: {e}")

    def save_settings(self, *args):
        store = JsonStore(os.path.join(BASE_DIR, 'settings.json'))
        store.put('engineering_specs',
                  kz1=float(self.f_k_z1.text or 1.10),
                  kz2=float(self.f_k_z2.text or 1.07),
                  kz3=float(self.f_k_z3.text or 1.05),
                  v_zapas=float(self.f_v_zapas.text or 3.0),
                  v_crit_spo=float(self.f_v_crit_spo.text or 0.5))
        
        calc_scr = self.manager.get_screen('calc')
        if hasattr(calc_scr, 'refresh_ui_elements'):
            calc_scr.refresh_ui_elements()
            
        # Интерактивный отклик на кнопке сохранения
        self.btn_save_config.text = "КОНФИГУРАЦИЯ УСПЕШНО СОХРАНЕНА И ПРИМЕНЕНА!"
        self.btn_save_config.md_bg_color = (0.1, 0.5, 0.8, 1)
        
        def restore_btn(dt):
            self.btn_save_config.text = "СОХРАНИТЬ КОНФИГУРАЦИЮ"
            self.btn_save_config.md_bg_color = (0.1, 0.6, 0.2, 1)
        Clock.schedule_once(restore_btn, 2.5)
        print("Конфигурация успешно сохранена и применена!")



class CementAcidScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.v_pipe = 3.02 
        self.mode = "CEMENT"
        
        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))

        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(45),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        btn_back = MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: self.go_back()
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(text="ЦМОСТ / ВАННЫ", theme_text_color="Custom", 
                                       text_color=(1,1,1,1), font_style="H6"))
        main_layout.add_widget(self.toolbar)

        # 2. ОСНОВНОЙ КОНТЕЙНЕР СО СКРОЛЛОМ
        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(10), padding=dp(15))

        # ТАБЫ (ОСТАЮТСЯ СВЕРХУ)
        tab_grid = MDGridLayout(cols=2, spacing=dp(5), size_hint_y=None, height=dp(45))
        self.b1 = MDFillRoundFlatButton(text="ЦЕМЕНТ", size_hint=(1,1), on_release=lambda x: self.switch_mode("CEMENT"))
        self.b2 = MDFillRoundFlatButton(text="ВАННА", size_hint=(1,1), on_release=lambda x: self.switch_mode("ACID"))
        tab_grid.add_widget(self.b1); tab_grid.add_widget(self.b2)
        self.container.add_widget(tab_grid)

        # ПОЛЯ ВВОДА
        self.inputs_box = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(8))
        def create_f(hint, text=""):
            f = MDTextField(hint_text=hint, text=text, mode="fill", font_size="15sp", 
                               size_hint_y=None, height=dp(34), input_filter='float')
            f.bind(focus=lambda ins, val: Clock.schedule_once(lambda dt: ins.select_all(), 0.1) if val else None)
            return f

        self.f_diam = create_f("Ø экспл. колонны (мм)", "168")
        self.f_depth = create_f("Глубина воронки (м)", "1500")
        self.f_target = create_f("Интервал установки (м)", "1550")
        self.f_vol = create_f("Объем смеси (м³)", "1.5")
        
        for f in [self.f_diam, self.f_depth, self.f_target, self.f_vol]: 
            self.inputs_box.add_widget(f)
        self.container.add_widget(self.inputs_box)

        # ОКНО РЕЗУЛЬТАТА
        self.res_label = MDLabel(
            text="Введите данные и выберите НКТ", 
            theme_text_color="Custom", text_color=(1,1,1,1), 
            markup=True, size_hint_y=None, line_height=1.3
        )
        self.res_label.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1] + dp(20)))
        self.container.add_widget(self.res_label)

        # 3. НИЖНИЙ БЛОК КНОПОК (ВЫБОР НКТ И РАССЧИТАТЬ)
        # Оборачиваем в AnchorLayout для принудительной растяжки на весь экран
        
        # КНОПКА ВЫБОРА НКТ
        anchor_nkt = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(52))
        self.btn_nkt = MDFillRoundFlatButton(
            text="ВЫБРАТЬ НКТ (л/м)", 
            size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), 
            on_release=self.open_nkt_menu
        )
        self.btn_nkt.size_hint_max_x = dp(2000)
        anchor_nkt.add_widget(self.btn_nkt)
        self.container.add_widget(anchor_nkt)

        # КНОПКА РАСЧЕТА
        anchor_calc = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(55))
        self.btn_calc = MDFillRoundFlatIconButton(
            text="РАССЧИТАТЬ ТЕХНОЛОГИЮ", icon="calculator", 
            size_hint=(1, 1), md_bg_color=(0.1, 0.6, 0.2, 1), 
            on_release=self.do_calc
        )
        self.btn_calc.size_hint_max_x = dp(2000)
        anchor_calc.add_widget(self.btn_calc)
        self.container.add_widget(anchor_calc)

        scroll.add_widget(self.container)
        main_layout.add_widget(scroll)
        self.add_widget(main_layout)
        self.switch_mode("CEMENT")

    def switch_mode(self, mode):
        self.mode = mode
        active, inactive = (0.1, 0.4, 0.8, 1), (0.2, 0.2, 0.2, 1)
        self.b1.md_bg_color = active if mode == "CEMENT" else inactive
        self.b2.md_bg_color = active if mode == "ACID" else inactive
        self.f_vol.hint_text = "Объем цемента (м³)" if mode == "CEMENT" else "Объем кислоты (м³)"

    def open_nkt_menu(self, b):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True); ws = wb.active
            headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            rows = [dict(zip(headers, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]
            items = [{"viewclass":"OneLineListItem","text":f"НКТ {r['Тип_НКТ']}","on_release":lambda x=r:self.set_nkt(x)} for r in rows]
            self.menu = MDDropdownMenu(caller=b, items=items, width_mult=4); self.menu.open()
        except: self.res_label.text = "[color=#FF5555]Ошибка базы Excel[/color]"

    def set_nkt(self, r):
        d_in = float(r.get('Внутр_d_мм', 62))
        self.v_pipe = (d_in**2) * 0.0007854 * 1000
        self.btn_nkt.text = f"НКТ {r['Тип_НКТ']} ({self.v_pipe:.2f} л/м)"
        self.menu.dismiss()

    def do_calc(self, *args):
        try:
            def val(f): return float(f.text or 0)
            D, Hv, H_tar, V_sm = val(self.f_diam), val(self.f_depth), val(self.f_target), val(self.f_vol)
            
            # Геометрия (л/м)
            v_col_litr = ((D-15)**2) * 0.0007854 * 1000
            # Объем НКТ (сколько влезет в трубы до низа)
            v_nkt_total = (self.v_pipe * Hv) / 1000
            # Объем продавки (сколько качать в трубы, чтобы вытолкнуть смесь ниже воронки)
            h_diff = max(0, H_tar - Hv)
            v_prodavka = (v_col_litr * h_diff) / 1000

            if self.mode == "CEMENT":
                bags = int((V_sm * 1000) / 36)
                res = (
                    f"[size=18sp][color=#33FF33]ИНСТРУКЦИЯ: ЦЕМЕНТНЫЙ МОСТ[/color][/size]\n\n"
                    f"1. [b]Закачка:[/b] Закачай в трубы [b]{V_sm} м³[/b] цемента.\n"
                    f"   (Это примерно [b]{bags} мешков[/b] по 50кг).\n"
                    f"2. [b]Продавка:[/b] Качай следом воду/раствор в объеме [b]{v_nkt_total:.2f} м³[/b].\n"
                    f"   [i]*Важно: Не больше и не меньше, чтобы цемент вышел из труб.*[/i]\n"
                    f"3. [b]Подъем:[/b] Сразу подними трубы выше [b]{H_tar - 50} м[/b].\n"
                    f"4. [b]Срезка:[/b] Качай раствор в затруб (обратка), чтобы вымыть лишний цемент."
                )
            else:
                res = (
                    f"[size=18sp][color=#33FF33]ИНСТРУКЦИЯ: КИСЛОТНАЯ ВАННА[/color][/size]\n\n"
                    f"1. [b]Закачка:[/b] Закачай в трубы [b]{V_sm} м³[/b] кислоты.\n"
                    f"2. [b]Доводка:[/b] Качай раствор в трубы [b]{v_nkt_total:.2f} м³[/b].\n"
                    f"   [b]ЗАТРУБ ДОЛЖЕН БЫТЬ ОТКРЫТ![/b] (Ждем выхода воздуха/жидкости).\n"
                    f"3. [b]Продавка:[/b] Как только докачал объем выше, [color=#FF5555][b]ЗАКРОЙ ЗАТРУБ![/b][/color]\n"
                    f"4. [b]Давим в пласт:[/b] Качай еще [b]{v_prodavka:.2f} м³[/b] в трубы.\n"
                    f"   (Кислота уйдет из под воронки прямо в зону перфорации).\n"
                    f"5. [b]Финиш:[/b] Закрой задвижки на трубах. Ожидание реакции 2-4 часа."
                )
            self.res_label.text = res
        except: 
            self.res_label.text = "[color=#FF5555]ОШИБКА: Проверь, все ли цифры ввел![/color]"

class CalculatorHubScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.selected_weight_m = 9.5
        self.pipe_name = "НКТ 73"
        self.tank_shape = "RECT"

        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))
        
        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(50),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        from kivymd.uix.button import MDIconButton
        btn_back = MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: setattr(self.manager, 'current', 'menu')
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(text="ИНЖЕНЕРНЫЕ РАСЧЕТЫ", theme_text_color="Custom", 
                                       text_color=(1,1,1,1), font_style="H6"))
        main_layout.add_widget(self.toolbar)

        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(15), padding=dp(20))

        # --- РАЗДЕЛ 1: РАСЧЕТ ВЕСА (ИВЭ) ---
        self.container.add_widget(MDLabel(text="ВЕС ИНСТРУМЕНТА С ПЛАВУЧЕСТЬЮ:", font_style="Button", theme_text_color="Hint"))
        
        # Кнопка выбора трубы на всю ширину
        anchor_pipe = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(50))
        self.btn_pipe = MDFillRoundFlatButton(text=f"ТРУБА: {self.pipe_name}", size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), on_release=self.menu_pipes)
        self.btn_pipe.size_hint_max_x = dp(2000)
        anchor_pipe.add_widget(self.btn_pipe)
        self.container.add_widget(anchor_pipe)

        self.f_depth = self.create_small_f("Глубина спуска (м)")
        self.f_rho_fluid = self.create_small_f("Плотность раствора (г/см³)", "1.00")
        self.container.add_widget(self.f_depth)
        self.container.add_widget(self.f_rho_fluid)

        # ТАБЛО ИВЭ (Крупно, симметрично)
        self.ive_display = MDGridLayout(
            cols=2, size_hint_y=None, height=dp(85), 
            padding=[dp(10), dp(10)], spacing=dp(10), md_bg_color=(0, 0, 0, 1)
        )
        
        # Блок "В ВОЗДУХЕ"
        box_v = MDBoxLayout(orientation='vertical', spacing=dp(2))
        box_v.add_widget(MDLabel(text="В ВОЗДУХЕ", halign="center", font_style="Caption", theme_text_color="Hint"))
        self.res_air_val = MDLabel(text="0.0", halign="center", font_style="H5", theme_text_color="Custom", text_color=(1,1,1,1), bold=True)
        box_v.add_widget(self.res_air_val)
        
        # Блок "НА КРЮКЕ"
        box_k = MDBoxLayout(orientation='vertical', spacing=dp(2))
        box_k.add_widget(MDLabel(text="НА КРЮКЕ", halign="center", font_style="Caption", theme_text_color="Hint"))
        self.res_hook_val = MDLabel(text="0.0", halign="center", font_style="H4", theme_text_color="Custom", text_color=(0.2, 1, 0.2, 1), bold=True)
        box_k.add_widget(self.res_hook_val)
        
        self.ive_display.add_widget(box_v)
        self.ive_display.add_widget(box_k)
        self.container.add_widget(self.ive_display)
        
        # Кнопка расчета на всю ширину
        anchor_calc_w = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(52))
        btn_calc_w = MDFillRoundFlatIconButton(text="РАССЧИТАТЬ ВЕС", icon="weight-lifter", size_hint=(1, 1), on_release=self.calc_weight)
        btn_calc_w.size_hint_max_x = dp(2000)
        anchor_calc_w.add_widget(btn_calc_w)
        self.container.add_widget(anchor_calc_w)

        # --- РАЗДЕЛ 2: СМЕШИВАНИЕ ---
        self.container.add_widget(self.add_line())
        self.container.add_widget(MDLabel(text="СМЕШИВАНИЕ ДВУХ РАСТВОРОВ:", font_style="Button", theme_text_color="Hint"))
        
        self.f_v1 = self.create_small_f("Объем 1 (м³)"); self.f_r1 = self.create_small_f("Ро 1 (г/см³)")
        self.f_v2 = self.create_small_f("Объем 2 (м³)"); self.f_r2 = self.create_small_f("Ро 2 (г/см³)")
        self.container.add_widget(self.f_v1); self.container.add_widget(self.f_r1)
        self.container.add_widget(self.f_v2); self.container.add_widget(self.f_r2)
        
        anchor_mix = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(50))
        btn_mix = MDFillRoundFlatButton(text="РАССЧИТАТЬ СМЕСЬ", size_hint=(1, 1), md_bg_color=(0.1, 0.6, 0.2, 1), on_release=self.calc_mixing)
        btn_mix.size_hint_max_x = dp(2000)
        anchor_mix.add_widget(btn_mix)
        self.container.add_widget(anchor_mix)

        self.res_mix = MDLabel(text="Итог: введите данные", halign="left", theme_text_color="Custom", text_color=(1,1,1,1), markup=True)
        self.container.add_widget(self.res_mix)

        # --- РАЗДЕЛ 3: ОБЪЕМ ЕМКОСТЕЙ ---
        self.container.add_widget(self.add_line())
        self.container.add_widget(MDLabel(text="ОБЪЕМ ЖИДКОСТИ В ЕМКОСТИ:", font_style="Button", theme_text_color="Hint"))
        
        shape_grid = MDBoxLayout(orientation='horizontal', spacing=dp(5), size_hint_y=None, height=dp(45))
        self.btn_rect = MDFillRoundFlatButton(text="ПРЯМОУГОЛЬНАЯ", size_hint=(1,1), on_release=lambda x: self.set_shape("RECT"))
        self.btn_cyl = MDFillRoundFlatButton(text="ЦИЛИНДР (ГОР.)", size_hint=(1,1), on_release=lambda x: self.set_shape("CYL"))
        shape_grid.add_widget(self.btn_rect); shape_grid.add_widget(self.btn_cyl)
        self.container.add_widget(shape_grid)

        self.f_t_l = self.create_small_f("Длина емкости (м)")
        self.f_t_w_d = self.create_small_f("Ширина емкости (м)")
        self.f_t_h = self.create_small_f("Уровень взлива (м)")
        self.container.add_widget(self.f_t_l); self.container.add_widget(self.f_t_w_d); self.container.add_widget(self.f_t_h)
        
        anchor_tank = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(50))
        btn_tank = MDFillRoundFlatButton(text="ПОСЧИТАТЬ ОБЪЕМ", size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), on_release=self.calc_tank_vol)
        btn_tank.size_hint_max_x = dp(2000)
        anchor_tank.add_widget(btn_tank)
        self.container.add_widget(anchor_tank)

        self.res_tank = MDLabel(text="Объем: --- м³", halign="left", bold=True)
        self.container.add_widget(self.res_tank)

        self.container.add_widget(MDBoxLayout(size_hint_y=None, height=dp(40))) 
        scroll.add_widget(self.container);
        main_layout.add_widget(scroll); 
        self.add_widget(main_layout)
        self.set_shape("RECT")

    # --- ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ---
    def create_small_f(self, hint, text=""):
        f = MDTextField(hint_text=hint, text=text, mode="fill", size_hint_y=None, height=dp(40), input_filter='float')
        f.bind(focus=lambda ins, val: Clock.schedule_once(lambda dt: ins.select_all(), 0.1) if val else None)
        return f

    def add_line(self):
        return MDBoxLayout(size_hint_y=None, height=dp(1), md_bg_color=(0.3, 0.3, 0.3, 1))

    def set_shape(self, shape):
        self.tank_shape = shape
        active, inactive = (0.1, 0.4, 0.8, 1), (0.2, 0.2, 0.2, 1)
        self.btn_rect.md_bg_color = active if shape == "RECT" else inactive
        self.btn_cyl.md_bg_color = active if shape == "CYL" else inactive
        self.f_t_w_d.hint_text = "Ширина (м)" if shape == "RECT" else "Диаметр (м)"

    def menu_pipes(self, b):
        # Список кортежей: (Отображаемый текст, Чистый вес для расчетов)
        list_p = [
            ("НКТ 60 (7.1 кг/м)", 7.1), 
            ("НКТ 73 (9.5 кг/м)", 9.5), 
            ("НКТ 89 (13.5 кг/м)", 13.5), 
            ("СБТ 73 (14.5 кг/м)", 14.5)
        ]
        # Важно: передаем в lambda весь кортеж x
        items = [{
            "viewclass": "OneLineListItem", 
            "text": i[0], 
            "on_release": lambda x=i: self.set_pipe(x)
        } for i in list_p]
        
        self.m_p = MDDropdownMenu(caller=b, items=items, width_mult=4)
        self.m_p.open()

    def set_pipe(self, val):
        # val[0] - это текст (напр. "НКТ 73 (9.5 кг/м)")
        # val[1] - это число (9.5)
        self.pipe_name = val[0].split(" (")[0]
        self.selected_weight_m = val[1]
        
        self.btn_pipe.text = f"ТРУБА: {self.pipe_name}"
        self.m_p.dismiss()
        # Сразу пересчитываем, если глубина уже введена
        self.calc_weight()

    def calc_weight(self, *args):
        try:
            L = float(self.f_depth.text or 0)
            rho_f = float(self.f_rho_fluid.text or 1.0)
            w_air = (L * self.selected_weight_m) / 1000
            w_hook = w_air * (1 - rho_f / 7.85)
            self.res_air_val.text = f"{w_air:.1f} т"
            self.res_hook_val.text = f"{w_hook:.1f} т"
            # Цвет красный при превышении 60 тонн
            self.res_hook_val.text_color = (1, 0.2, 0.2, 1) if w_hook > 60 else (0.2, 1, 0.2, 1)
        except: pass

    def calc_mixing(self, *args):
        try:
            v1, r1 = float(self.f_v1.text or 0), float(self.f_r1.text or 0)
            v2, r2 = float(self.f_v2.text or 0), float(self.f_r2.text or 0)
            if (v1 + v2) > 0:
                res = (v1 * r1 + v2 * r2) / (v1 + v2)
                self.res_mix.text = f"• Плотность смеси: [b]{res:.2f} г/см³[/b]\n• Общий объем: [b]{v1+v2:.1f} м³[/b]"
        except: self.res_mix.text = "Ошибка данных"

    def calc_tank_vol(self, *args):
        try:
            L, WD, H = float(self.f_t_l.text or 0), float(self.f_t_w_d.text or 0), float(self.f_t_h.text or 0)
            if self.tank_shape == "RECT":
                vol = L * WD * H
            else:
                vol = (math.pi * (WD/2)**2 * L) * (H/WD)
            self.res_tank.text = f"Объем жидкости: [b]{vol:.2f} м³[/b]"
        except: pass

class GnvpScreen(BaseScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.method = "DRILLER"
        self.v_vnutr_nkt = 3.02 
        
        main_layout = MDBoxLayout(orientation='vertical', size_hint=(1, 1))
        
        # 1. ШАПКА (Оранжевый — цвет опасности)
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(50),
            md_bg_color=(0.8, 0.4, 0, 1), padding=[dp(15), 0, dp(10), 0]
        )
        from kivymd.uix.button import MDIconButton
        btn_back = MDIconButton(
            icon="arrow-left", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: setattr(self.manager, 'current', 'menu')
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(text="ЛИСТ ГЛУШЕНИЯ (ГНВП)", theme_text_color="Custom", 
                                       text_color=(1,1,1,1), font_style="H6"))
        main_layout.add_widget(self.toolbar)

        scroll = ScrollView()
        self.container = MDBoxLayout(orientation='vertical', adaptive_height=True, spacing=dp(10), padding=dp(20))

        # 2. ИНТЕРАКТИВНЫЕ ОКНА (ПОЛЯ ВВОДА)
        def f_make(hint, val=""):
            f = MDTextField(hint_text=hint, text=val, mode="fill", input_filter='float', size_hint_y=None, height=dp(34))
            f.bind(focus=lambda ins, v: Clock.schedule_once(lambda dt: ins.select_all(), 0.1) if v else None)
            return f

        self.f_d_col = f_make("Ø колонны (мм)", "168")
        self.f_ckod = f_make("ЦКОД / Иск. забой (м)", "2500")
        self.f_h_hvost = f_make("Длина хвостовика 114 (м)", "0")
        self.f_h_v = f_make("Глубина воронки (м)", "2000")
        self.f_h_tvd = f_make("Глубина по вертикали (м)", "1800")
        self.f_sidpp = f_make("Р-трубное SIDPP (атм)", "0")
        self.f_sicp = f_make("Р-затрубное SICP (атм)", "15")
        self.f_rho_old = f_make("Ро раствора в скв. (г/см³)", "1.10")

        fields = [self.f_d_col, self.f_ckod, self.f_h_hvost, self.f_h_v, 
                  self.f_h_tvd, self.f_sidpp, self.f_sicp, self.f_rho_old]
        for f in fields: 
            self.container.add_widget(f)

        # 3. ВЫБОР НКТ (РАСПОЛОЖЕН ПОД ОКНАМИ)
        anchor_nkt = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(50))
        self.btn_nkt = MDFillRoundFlatButton(
            text="ВЫБРАТЬ ТИП НКТ ПОДВЕСКИ", 
            size_hint=(1, 1), md_bg_color=(0.1, 0.3, 0.6, 1), 
            on_release=self.open_nkt_menu
        )
        self.btn_nkt.size_hint_max_x = dp(2000)
        anchor_nkt.add_widget(self.btn_nkt)
        self.container.add_widget(anchor_nkt)

        # 4. ВЫБОР ТАКТИКИ (КНОПКИ НА ШИРИНУ ЭКРАНА)
        self.container.add_widget(MDLabel(text="ТАКТИКА ГЛУШЕНИЯ:", font_style="Caption", theme_text_color="Hint"))
        
        def create_wide_mode_btn(text, mode):
            anchor = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(45))
            btn = MDFillRoundFlatButton(text=text, size_hint=(1, 1), on_release=lambda x: self.set_m(mode))
            btn.size_hint_max_x = dp(2000)
            anchor.add_widget(btn)
            return anchor, btn

        self.anc1, self.m1 = create_wide_mode_btn("МЕТОД БУРИЛЬЩИКА (2 ЦИКЛА)", "DRILLER")
        self.anc2, self.m2 = create_wide_mode_btn("ОЖИДАНИЕ И УТЯЖЕЛЕНИЕ (1 ЦИКЛ)", "WAIT")
        self.anc3, self.m3 = create_wide_mode_btn("ЗАДАВКА (BULLHEADING)", "BULL")
        
        for a in [self.anc1, self.anc2, self.anc3]: self.container.add_widget(a)

        # 5. КНОПКА РАСЧЕТА (НА ВСЮ ШИРИНУ)
        anchor_calc = MDAnchorLayout(anchor_x='center', size_hint_y=None, height=dp(55))
        self.btn_calc = MDFillRoundFlatIconButton(
            text="РАССЧИТАТЬ ИНСТРУКЦИЮ", icon="alert-octagon", 
            size_hint=(1, 1), md_bg_color=(0.8, 0, 0, 1), 
            on_release=self.calculate
        )
        self.btn_calc.size_hint_max_x = dp(2000)
        anchor_calc.add_widget(self.btn_calc)
        self.container.add_widget(anchor_calc)

        # 6. ВЫВОД РЕЗУЛЬТАТОВ
        self.res = MDLabel(text="Введите данные замеров ГК", markup=True, halign="center", 
                          theme_text_color="Custom", text_color=(1,1,1,1), size_hint_y=None)
        self.res.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1] + dp(20)))
        self.container.add_widget(self.res)

        scroll.add_widget(self.container)
        main_layout.add_widget(scroll)
        self.add_widget(main_layout)
        self.set_m("DRILLER")

    def set_m(self, mode):
        self.method = mode
        active, inactive = (0.1, 0.4, 0.8, 1), (0.2, 0.2, 0.2, 1)
        self.m1.md_bg_color = active if mode == "DRILLER" else inactive
        self.m2.md_bg_color = active if mode == "WAIT" else inactive
        self.m3.md_bg_color = active if mode == "BULL" else inactive

    def open_nkt_menu(self, b):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True); ws = wb.active
            head = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            data = [dict(zip(head, r)) for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
            items = [{"viewclass": "OneLineListItem", "text": f"НКТ {r['Тип_НКТ']}", "on_release": lambda x=r: self.set_nkt(x)} for r in data]
            self.m_nkt = MDDropdownMenu(caller=b, items=items, width_mult=4, background_color=(0.1, 0.1, 0.15, 1))
            self.m_nkt.open()
        except: self.res.text = "[color=#FF5555]Ошибка базы данных[/color]"

    def set_nkt(self, r):
        d_in = float(r.get('Внутр_d_мм', 62))
        self.v_vnutr_nkt = (d_in**2) * 0.0007854 * 1000
        self.btn_nkt.text = f"НКТ {r['Тип_НКТ']} ({self.v_vnutr_nkt:.2f} л/м)"
        self.m_nkt.dismiss()

    def calculate(self, *args):
        try:
            def val(f): return float(f.text or 0)
            D, L_ckod, L_hvo, Hv = val(self.f_d_col), val(self.f_ckod), val(self.f_h_hvost), val(self.f_h_v)
            H_tvd, pt, pz, ro = val(self.f_h_tvd), val(self.f_sidpp), val(self.f_sicp), val(self.f_rho_old)
            
            h_ref = H_tvd if H_tvd > 0 else L_ckod
            if h_ref == 0: 
                self.res.text = "[color=#FFCC00]Укажите глубину![/color]"; return
            
            # 1. РАСЧЕТ ПЛОТНОСТИ (Запас +0.05 г/см3 по ПБ 08-624-03)
            # Если трубное давление (pt) равно 0, значит скважина стоит под давлением затруба (pz)
            p_izb = pt if pt > 0 else pz
            ro_k = round(ro + (p_izb * 10.197 / h_ref) + 0.05, 2)
            p_plas = int(h_ref * (ro + (pt * 10.197 / h_ref)) / 10.197)

            # 2. ГЕОМЕТРИЯ (литры на метр)
            v_vnutr_litr = self.v_vnutr_nkt  # Объем НКТ л/м
            v_col_litr = ((D - 15)**2) * 0.0007854 * 1000 # Объем колонны л/м
            
            # Объемы в м3
            v_nkt = round((v_vnutr_litr * Hv) / 1000, 2)
            v_annulus = round(((v_col_litr * Hv) / 1000) - (v_nkt * 0.35), 2) # Затруб
            v_pod = round((v_col_litr * (L_ckod - Hv)) / 1000, 2) # Под воронкой
            v_well_total = round(v_nkt + v_annulus + v_pod, 2)

            # --- ИНСТРУКЦИИ ПО РЕГЛАМЕНТУ ---
            if self.method == "DRILLER":
                tech = (
                    f"[color=#33FF33][b]МЕТОД БУРИЛЬЩИКА (2 ЦИКЛА):[/b][/color]\n"
                    f"• [b]1-й цикл:[/b] Вымыв пачки старым раствором.\n"
                    f"  Качай в затруб. Держи Р_затр = [b]{pz} атм[/b].\n"
                    f"  Пачка выйдет на устье через [color=#FFFF33][b]{v_nkt} м³[/b][/color] (объем НКТ).\n"
                    f"• [b]2-й цикл:[/b] Замещение на Ро = [b]{ro_k}[/b].\n"
                    f"  Качай в затруб. Давление Р_затр снижать до 0."
                )
            elif self.method == "WAIT":
                tech = (
                    f"[color=#33FF33][b]ОЖИДАНИЕ И УТЯЖЕЛЕНИЕ:[/b][/color]\n"
                    f"1. Приготовь [b]{v_well_total} м³[/b] раствора Ро = [b]{ro_k}[/b].\n"
                    f"2. Закачка в НКТ. Снижай Р_труб с [b]{pt}[/b] до [b]0[/b] атм\n"
                    f"   пропорционально объему [b]{v_nkt} м³[/b].\n"
                    f"3. При выходе из воронки держи Р_затр = [b]{pz} атм[/b]."
                )
            else: # BULLHEADING
                tech = (
                    f"[color=#FF5555][b]ЗАДАВКА (BULLHEADING):[/b][/color]\n"
                    f"1. [b]Глушение без циркуляции.[/b] Закрой трубное.\n"
                    f"2. Качай Ро = [b]{ro_k}[/b] в затрубное пространство.\n"
                    f"3. [b]ВНИМАНИЕ:[/b] Контроль Р_заб. Не превышать Р_погл.\n"
                    f"   При наличии МГРП: Р_max = [b]Р_откр - 20 атм[/b]."
                )

            self.res.text = (
                f"[size=20sp]ПЛОТНОСТЬ ГЛУШЕНИЯ: [b]{ro_k}[/b][/size]\n"
                f"Пластовое давление: [b]{p_plas} атм[/b]\n"
                f"Общий объем скважины: [b]{v_well_total} м³[/b]\n"
                f"--------------------------------\n"
                f"{tech}\n"
                f"--------------------------------\n"
                f"[i]Контроль КВД в течение 2-х часов после закачки.[/i]"
            )
        except Exception as e: 
            self.res.text = f"[color=#FF5555]Ошибка в расчетах[/color]"

class KillResultScreen(BaseScreen):  #Результаты глушения. Показывает итоговые цифры и пошаговую тактическую инструкцию.
    def __init__(self, **kw):
        super().__init__(**kw)
        layout = MDBoxLayout(orientation='vertical')
        
        # 1. ШАПКА
        self.toolbar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(45),
            md_bg_color=(0.1, 0.4, 0.8, 1), padding=[dp(15), 0, dp(10), 0]
        )
        from kivymd.uix.button import MDIconButton
        btn_back = MDIconButton(
            icon="arrow-left", pos_hint={'center_y': .5},
            theme_text_color="Custom", text_color=(1, 1, 1, 1),
            on_release=lambda x: self.go_back()
        )
        self.toolbar.add_widget(btn_back)
        self.toolbar.add_widget(MDLabel(text="РЕЗУЛЬТАТЫ РАСЧЕТА", theme_text_color="Custom", 
                                       text_color=(1,1,1,1), font_style="H6"))
        layout.add_widget(self.toolbar)

        # 2. ОБЛАСТЬ ВЫВОДА
        scroll = ScrollView()
        self.result_box = MDBoxLayout(orientation='vertical', adaptive_height=True, padding=dp(20), spacing=dp(15))
        
        # КАРТОЧКА С ОСНОВНЫМИ ПАРАМЕТРАМИ
        self.res_card = MDCard(
            orientation='vertical', padding=dp(15), spacing=dp(5),
            size_hint=(1, None), adaptive_height=True,
            md_bg_color=(0.15, 0.2, 0.3, 1), radius=[dp(10)]
        )
        self.result_text = MDLabel(
            text="Данные рассчитываются...", theme_text_color="Custom", text_color=(1, 1, 1, 1),
            font_style="Subtitle1", size_hint_y=None, markup=True, line_height=1.4
        )
        self.result_text.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1]))
        self.res_card.add_widget(self.result_text)
        self.result_box.add_widget(self.res_card)

        # ИНСТРУКЦИЯ (Тактика действий)
        self.result_box.add_widget(MDLabel(text="ПОРЯДОК ДЕЙСТВИЙ:", font_style="Caption", theme_text_color="Hint"))
        
        self.instruction_text = MDLabel(
            text="Инструкция формируется...", theme_text_color="Primary",
            font_style="Body2", size_hint_y=None, markup=True, line_height=1.3
        )
        self.instruction_text.bind(texture_size=lambda ins, sz: setattr(ins, 'height', sz[1] + dp(10)))
        self.result_box.add_widget(self.instruction_text)

        # 3. КНОПКА ПЕРЕХОДА К СИМУЛЯЦИИ
        self.btn_view_scheme = MDFillRoundFlatIconButton(
            text="ВИЗУАЛИЗАЦИЯ СКВАЖИНЫ", icon="eye-outline",
            size_hint_x=None, width=dp(280), pos_hint={'center_x': 0.5},
            md_bg_color=(0.1, 0.6, 0.2, 1), on_release=self.go_to_scheme
        )
        self.result_box.add_widget(self.btn_view_scheme)

        scroll.add_widget(self.result_box)
        layout.add_widget(scroll)
        self.add_widget(layout)

    def go_to_scheme(self, x):
        self.manager.current = 'well_view'

    def go_back(self):
        # Возвращаемся в калькулятор для правки данных
        self.manager.current = 'calc'

class MenuCard(MDCard):
    def __init__(self, text, icon, color, screen_name, **kwargs):
        super().__init__(**kwargs)
        self.size_hint = (None, None)
        self.radius = [dp(15)]
        self.orientation = "vertical"
        self.md_bg_color = color
        self.ripple_behavior = True
        self.target = screen_name
        self.padding = [dp(5), dp(10), dp(5), dp(5)] # Увеличили верхний отступ
        self.spacing = dp(2) # Зазор между иконкой и текстом

        # Иконка (уменьшена до 32dp, чтобы не давить на текст)
        self.icon_widget = MDIconButton(
            icon=icon,
            theme_icon_color="Custom",
            icon_color=(1, 1, 1, 1),
            pos_hint={"center_x": .5},
            icon_size=dp(32), 
            size_hint=(1, 0.45)
        )
        self.icon_widget.disabled = True
        self.add_widget(self.icon_widget)

        # Текст (улучшенная читаемость и адаптивность)
        self.label_widget = MDLabel(
            text=text,
            halign="center",
            valign="top",
            theme_text_color="Custom",
            text_color=(1, 1, 1, 1),
            bold=True,
            font_style="Caption",
            size_hint=(1, 0.55),
            line_height=0.95
        )
        self.add_widget(self.label_widget)

    def on_release(self):
        app = MDApp.get_running_app()
        if hasattr(app, 'sm'):
            app.sm.current = self.target

class MainMenuScreen(MDScreen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.root_ui = FloatLayout()

        # 1. ФОН (Картинка вышки)
        self.root_ui.add_widget(FitImage(source=BACKGROUND_PATH, opacity=0.45))

        # 2. ЗАГОЛОВОК И СЛОГАН (ЦЕНТРИРОВАНИЕ)
        self.title_label = Label(
            text="OIL MATE", bold=True, color=(1, 1, 1, 1),
            outline_color=(0, 0, 0, 1), outline_width=2,
            size_hint=(None, None), 
            halign='center',
            pos_hint={'center_x': 0.5, 'top': 0.97} 
        )
        self.subtitle_label = Label(
            text="«Твой расчет — твоя безопасность»",
            color=(1, 1, 1, 1), outline_color=(0, 0, 0, 1), outline_width=1,
            size_hint=(None, None), 
            halign='center',
            pos_hint={'center_x': 0.5, 'top': 0.87}
        )

        # 3. ЛОГОТИП
        self.logo = Image(
            source=LOGO_PATH, size_hint=(None, None),
            pos_hint={'right': 0.98, 'top': 0.99},
            allow_stretch=True, keep_ratio=True
        )

        # 4. СЕТКА КНОПОК (РОВНО 8 ПЛИТОК ДЛЯ СИММЕТРИИ)
        self.main_content = MDBoxLayout(orientation='vertical')
        self.spacer = MDBoxLayout(size_hint_y=None)
        self.grid_anchor = FloatLayout(size_hint_y=None) 
        self.grid = MDGridLayout(cols=2, spacing=dp(12), adaptive_size=True, pos_hint={'center_x': 0.5, 'top': 1})
        
        # Фирменные полупрозрачные цвета (Альфа-канал изменен на 0.55 для красивого эффекта стекла)
        c_blue = (0.1, 0.4, 0.8, 0.55)    # Фирменный полупрозрачный синий
        c_orange = (0.8, 0.4, 0, 0.55)   # Фирменный полупрозрачный оранжевый
        c_grey = (0.2, 0.25, 0.3, 0.55)   # Строгий полупрозрачный серый (только для справочника/настроек)

        # Переставлены местами, ловильные перекрашены в синий, добавлена прозрачность
        grid_items = [
            ("ГЛУШЕНИЕ", "water-alert", c_blue, "calc"),
            ("КОНТРОЛЬ\nДОЛИВА", "water-pump", c_blue, "spo"),
            ("ВЗД ", "piston", c_blue, "vzd"),
            ("ЦМОСТ /\nВАННЫ", "format-color-fill", c_blue, "cmoist"),
            ("ЛОВИЛЬНЫЕ\nРАБОТЫ", "hook", c_blue, "fishing"),               # Перенесено выше и стало синим!
            ("ИНЖЕНЕРНЫЙ\nХАБ", "calculator", c_blue, "calc_hub"),
            ("ИНСТРУМЕНТ\n(СПРАВКА)", "book-open-variant", c_grey, "threads"), # Смещено ниже
            ("НАСТРОЙКИ\nПРИЛОЖЕНИЯ", "cog", c_grey, "settings")
        ]
        for t, i, c, s in grid_items:
            self.grid.add_widget(MenuCard(t, i, c, s))

        self.main_content.add_widget(self.spacer)
        self.grid_anchor.add_widget(self.grid)
        self.main_content.add_widget(self.grid_anchor)
        self.main_content.add_widget(MDBoxLayout()) 

        # 5. ПОДПИСЬ
        self.footer_label = MDLabel(
            text="URTUSHKOV © 2024", halign="center",
            theme_text_color="Hint", size_hint_y=None, height=dp(40)
        )
        self.main_content.add_widget(self.footer_label)

        self.root_ui.add_widget(self.title_label)
        self.root_ui.add_widget(self.subtitle_label)
        self.root_ui.add_widget(self.logo)
        self.root_ui.add_widget(self.main_content)
        self.add_widget(self.root_ui)

        Window.bind(on_resize=self.update_ui_sizes)
        self.update_ui_sizes()

    def update_ui_sizes(self, *args):
        w, h = Window.width, Window.height
        
        self.title_label.font_size = w * 0.11
        self.title_label.size = (w, self.title_label.font_size * 1.5)
        
        self.subtitle_label.font_size = w * 0.038
        self.subtitle_label.size = (w, self.subtitle_label.font_size * 2)
        
        logo_dim = min(w * 0.14, h * 0.08)
        self.logo.size = (logo_dim, logo_dim)
        
        self.spacer.height = h * 0.18
        self.grid_anchor.height = h * 0.65 

        # Пропорции плиток под сетку 2х4
        card_w = (w - dp(45)) / 2
        card_h = max(dp(95), min(h * 0.15, dp(135)))
        
        for card in self.grid.children:
            if isinstance(card, MenuCard):
                card.size = (card_w, card_h)
                card.label_widget.font_size = sp(card_h * 0.11)


# --- ГЛАВНОЕ ЯДРО ---
class KrsBookApp(MDApp):      #Ядро. Запускает всё приложение, регистрирует экраны, настраивает тему и управляет системной кнопкой «Назад».
    # Глобальная переменная для текстуры фона (чтобы не грузить её на каждом экране заново)
    global_texture = ObjectProperty(None)

    def build(self):
        # 1. Название приложения (должно совпадать с заголовком в buildozer.spec)
        self.title = "Помощник бурильщика ОС"
        
        # 2. Установка иконки
        if os.path.exists(ICON_PATH):
            self.icon = ICON_PATH
        
        # 3. Настройка темы KivyMD
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.material_style = "M3"
        
        # Настройка клавиатуры: при вводе текста экран сдвигается вверх
        Window.softinput_mode = "below_target" 

        # 4. Предзагрузка фона
        if os.path.exists(BACKGROUND_PATH):
            try:
                self.global_texture = CoreImage(BACKGROUND_PATH).texture
            except Exception as e:
                print(f"Ошибка загрузки фона: {e}")

        # 5. Инициализация менеджера экранов
        self.sm = ScreenManager(transition=FadeTransition())
        
        # Список всех экранов приложения
        screens = [
            LoadingScreen(name='loading'),
            MainMenuScreen(name='menu'),
            CalcScreen(name='calc'),
            KillResultScreen(name='kill_res'),
            SpoScreen(name='spo'),
            VzdScreen(name='vzd'),
            ThreadsScreen(name='threads'),
            CementAcidScreen(name='cmoist'),
            WellViewScreen(name='well_view'),
            FishingScreen(name='fishing'),
            CalculatorHubScreen(name='calc_hub'),
            SettingsScreen(name='settings') # <--- ОБЯЗАТЕЛЬНО ДОБАВЬ СЮДА
        ]


        # Добавление экранов в менеджер
        for screen in screens:
            self.sm.add_widget(screen)

        return self.sm

    def show_disclaimer(self):
        """Окно предупреждения (Отказ от ответственности)"""
        # Используем JsonStore для сохранения факта принятия условий
        self.store = JsonStore(os.path.join(BASE_DIR, 'settings.json'))
        
        if self.store.exists('disclaimer') and self.store.get('disclaimer')['accepted']:
            return

        disclaimer_text = (
            "Данное приложение является вспомогательным инженерным инструментом. "
            "Все расчеты носят справочный характер.\n\n"
            "Разработчик не несет ответственности за любые последствия (аварии, "
            "материальный ущерб), возникшие в результате использования данных.\n\n"
            "Сверяйте результаты с утвержденным Планом работ и РД."
        )

        self.dialog = MDDialog(
            title="ОТКАЗ ОТ ОТВЕТСТВЕННОСТИ",
            text=disclaimer_text,
            type="confirmation", 
            auto_dismiss=False,
            size_hint_x=0.85,    
            buttons=[
                MDFlatButton(
                    text="Я ПРИНИМАЮ",
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release=self.close_disclaimer
                ),
            ],
        )
        self.dialog.open()

    def close_disclaimer(self, *args):
        self.store.put('disclaimer', accepted=True)
        self.dialog.dismiss()

    def on_start(self):
        """События при старте приложения"""
        # Привязываем физическую или системную кнопку 'Назад'
        from kivy.base import EventLoop
        EventLoop.window.bind(on_keyboard=self.hook_keyboard)

    def hook_keyboard(self, window, key, *args):
        # Код клавиши 'Назад' (ESC) = 27
        if key == 27:  
            current = self.sm.current
            if current in ['loading', 'menu']:
                return False  # Закрыть приложение
            elif current == 'well_view':
                self.sm.current = 'kill_res'
                return True
            elif current == 'kill_res':
                self.sm.current = 'calc'
                return True
            else:
                self.sm.current = 'menu'
                return True
        return False

# --- ТОЧКА ВХОДА ---
if __name__ == "__main__":
    KrsBookApp().run()
