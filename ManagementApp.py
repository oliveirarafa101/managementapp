from kivy.app import App
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.button import Button
from kivy.uix.image import Image
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.metrics import dp
import openpyxl

class Root(App):
    def build(self):
        layout = FloatLayout()
        
        # Adiciona a imagem de um caminhão como plano de fundo
        truck_image = Image(source='truck.png', allow_stretch=True, keep_ratio=False)
        layout.add_widget(truck_image)
        
        # Adiciona campos de entrada na frente da imagem, centralizados horizontalmente
        self.inputs = {}
        y_pos = 0.9  # Define a posição vertical inicial das caixas de entrada
        for idx, header in enumerate(["CLIENTE A FATURAR", "CLIENTE", "DATA", "CONTENTOR", "VIATURA", "GUIA TRANSPORTE", "MOTORISTA"]):
            input_field = TextInput(hint_text=header, size_hint=(0.5, 0.05), pos_hint={'center_x': 0.5, 'top': y_pos}, background_color=(0.8, 0.8, 0.8, 1))
            self.inputs[header] = input_field
            layout.add_widget(input_field)
            y_pos -= 0.09  # Atualiza a posição vertical para o próximo campo
        
        # Adiciona botão de salvar
        save_button = Button(text="Save to Excel", size_hint=(0.3, 0.1), pos_hint={'center_x': 0.3, 'y': 0.05})
        save_button.bind(on_press=self.save_to_excel)
        layout.add_widget(save_button)

        # Adiciona botão para mostrar os dados já inseridos
        show_data_button = Button(text="Show Inserted Data", size_hint=(0.3, 0.1), pos_hint={'center_x': 0.7, 'y': 0.05})
        show_data_button.bind(on_press=self.show_inserted_data)
        layout.add_widget(show_data_button)
        
        return layout

    def save_to_excel(self, instance):
        file_path = "dados_camioes.xlsx"
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # Cabeçalhos
        headers = ["CLIENTE A FATURAR", "CLIENTE", "DATA", "CONTENTOR", "VIATURA", "GUIA TRANSPORTE", "MOTORISTA"]
        if sheet.max_row == 1:
            sheet.append(headers)

        # Obtém os dados dos campos de entrada
        data = [self.inputs[header].text for header in headers]
        sheet.append(data)

        # Salva o arquivo Excel
        workbook.save(file_path)

        # Feedback ao usuário
        success_popup = Popup(title='Success', content=Button(text='Data saved successfully!'), size_hint=(None, None), size=(400, 400))
        success_popup.open()

    def show_inserted_data(self, instance):
        file_path = "dados_camioes.xlsx"
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # Mostra os dados já inseridos em uma tabela organizada
        data_grid = GridLayout(cols=len(sheet[1]), spacing=5, size_hint_y=None, pos_hint={'center_x': 0.5, 'y': 0.2})
        data_grid.bind(minimum_height=data_grid.setter('height'))

        # Adiciona cabeçalhos
        for cell in sheet[1]:
            label = Label(text=str(cell.value), size_hint=(None, None), size=(dp(100), dp(40)), color=(1, 1, 1, 1))
            data_grid.add_widget(label)

        # Adiciona dados
        for row in sheet.iter_rows(min_row=2, values_only=True):
            for item in row:
                label = Label(text=str(item), size_hint=(None, None), size=(dp(100), dp(40)), color=(1, 1, 1, 1))
                data_grid.add_widget(label)

        data_popup = Popup(title='', content=data_grid, size_hint=(None, None), size=(dp(800), dp(400)))
        data_popup.open()

if __name__ == "__main__":
    Root().run()
