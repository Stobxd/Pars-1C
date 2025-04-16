import flet as ft

from flet import (
    ElevatedButton,
    FilePicker,
    FilePickerResultEvent,
    Page,
    Row,
    Text,
    icons,
)
def main(page: ft.Page):

    main_view = ft.Tabs(
        selected_index=0,
        animation_duration=600,

        tabs=[
            
            ft.Tab(text="1С Буратино", #Заголовок табки
                   
                ontent = ft.Container(
                    content=ft.Column(
                        [
                            ft.Row([ft.TextButton("Buy tickets"), ft.TextButton("Listen")], alignment=ft.MainAxisAlignment.END)
,
                        ]
                    )






                )
            )
        ],


    )

    page.add(main_view)
ft.app(main)