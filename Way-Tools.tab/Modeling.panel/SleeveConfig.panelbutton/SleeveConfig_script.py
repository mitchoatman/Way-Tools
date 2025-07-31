import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('0.5')

with open(filepath, 'r') as f:
    SleeveLength = float(f.read()) * 12

# WPF Form (no XAML)
class SleeveForm(Window):
    def __init__(self, initial_value):
        self.Title = "Sleeve Configuration"
        self.Width = 300
        self.Height = 160
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result_value = None

        grid = Grid()
        grid.Margin = Thickness(10)

        for _ in range(3):
            grid.RowDefinitions.Add(RowDefinition())

        self.Content = grid

        label = Label()
        label.Content = "Default Sleeve Length (Inches):"
        label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        self.textbox = TextBox()
        self.textbox.Text = str(round(initial_value, 3))
        self.textbox.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.textbox, 1)
        grid.Children.Add(self.textbox)

        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.ok_clicked
        Grid.SetRow(ok_button, 2)
        grid.Children.Add(ok_button)

        self.textbox.Focus()
        self.textbox.SelectAll()

    def ok_clicked(self, sender, args):
        try:
            self.result_value = float(self.textbox.Text)
            self.DialogResult = True
            self.Close()
        except:
            self.textbox.Text = "0.5"

form = SleeveForm(SleeveLength)
updated_length = None
if form.ShowDialog() and form.DialogResult:
    updated_length = form.result_value / 12
    with open(filepath, 'w') as f:
        f.write(str(updated_length))
