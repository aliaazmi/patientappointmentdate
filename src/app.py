import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from dash.dependencies import Output, Input, State
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle

app = dash.Dash(_name_, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server
app.title = 'Patient Appointment Date'

app.layout = dbc.Container(
    [
        html.H1('Patient Appointment Date', className='text-center mt-4 mb-4'),
        dbc.FormGroup(
            [
                dbc.Label('Enter your name:', className='mr-2'),
                dbc.Input(id='name-input', type='text'),
            ],
            className='mb-3',
        ),
        dbc.FormGroup(
            [
                dbc.Label('Select Appointment Type:', className='mr-2'),
                dcc.Dropdown(
                    id='appointment-type-dropdown',
                    options=[
                        {'label': 'Imaging', 'value': 'imaging'},
                        {'label': 'Cycle Visit', 'value': 'cycle_visit'}
                    ],
                    value='imaging',
                    clearable=False,
                    style={'width': '50%'}
                ),
            ],
            className='mb-3',
        ),
        dbc.FormGroup(
            [
                dbc.Label('Date of C1D1 or Randomization (YYYY-MM-DD):', className='mr-2'),
                dbc.Input(id='start-date-input', type='text'),
            ],
            className='mb-3',
        ),
        dbc.FormGroup(
            [
                dbc.Label('Select Scheduling Interval:', className='mr-2'),
                dcc.Dropdown(
                    id='interval-dropdown',
                    options=[
                        {'label': 'Day', 'value': 'day'},
                        {'label': 'Week', 'value': 'week'},
                        {'label': 'Month', 'value': 'month'}
                    ],
                    value='day',
                    clearable=False,
                    style={'width': '30%'}
                ),
            ],
            className='mb-3',
        ),
        dbc.FormGroup(
            [
                dbc.Label('Enter the number of intervals between appointments:', className='mr-2'),
                dbc.Input(id='interval-input', type='number'),
            ],
            className='mb-3',
        ),
        dbc.FormGroup(
            [
                dbc.Label('Enter the number of plus/minus days:', className='mr-2'),
                dbc.Input(id='plus-minus-days-input', type='number'),
            ],
            className='mb-3',
        ),
        dbc.Button('Calculate Appointments', id='calculate-button', n_clicks=0, className='mb-3'),
        html.Div(id='appointment-table-div'),
        dbc.Alert(
            id='alert',
            is_open=False,
            duration=4000,
        ),
    ],
    className='p-5',
)


def calculate_appointment_dates(start_date, interval_type, interval_value, plus_minus_days):
    num_appointments = 30
    appointment_dates = []
    current_date = datetime.strptime(start_date, '%Y-%m-%d')
    interval_value = int(interval_value)
    plus_minus_days = int(plus_minus_days) if plus_minus_days else 0

    for i in range(num_appointments):
        appointment_dates.append(current_date.strftime('%Y-%m-%d'))

        plus_date = current_date + timedelta(days=plus_minus_days)
        appointment_dates.append(plus_date.strftime('%Y-%m-%d'))

        minus_date = current_date - timedelta(days=plus_minus_days)
        appointment_dates.append(minus_date.strftime('%Y-%m-%d'))

        if interval_type == 'day':
            current_date += timedelta(days=interval_value)
        elif interval_type == 'week':
            current_date += timedelta(weeks=interval_value)
        elif interval_type == 'month':
            current_date += relativedelta(months=interval_value)

    return appointment_dates


def generate_pdf_report(name, appointment_dates, appointment_type, plus_minus_days):
    doc = SimpleDocTemplate("appointment_report.pdf", pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    title = Paragraph("Patient Appointment Dates", styles['Title'])
    elements.append(title)

    name_paragraph = Paragraph(f"Patient Name: {name}", styles['Normal'])
    elements.append(name_paragraph)

    type_paragraph = Paragraph(f"Appointment Type: {appointment_type}", styles['Normal'])
    elements.append(type_paragraph)

    data = [['Cycle Number', 'Date', 'Plus Day Date', 'Minus Day Date']]
    for i, date in enumerate(appointment_dates):
        if i % 3 == 0:
            data.append([str(i // 3 + 1), date, '', ''])
        elif i % 3 == 1:
            data[-1][2] = date
        else:
            data[-1][3] = date

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), '#CCCCCC'),
        ('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), '#FFFFFF'),
        ('BOX', (0, 0), (-1, -1), 1, '#000000'),
        ('GRID', (0, 0), (-1, -1), 0.5, '#000000'),
    ]))
    elements.append(table)

    doc.build(elements)


@app.callback(
    Output('appointment-table-div', 'children'),
    Output('alert', 'is_open'),
    [Input('calculate-button', 'n_clicks')],
    [State('name-input', 'value'),
     State('start-date-input', 'value'),
     State('interval-dropdown', 'value'),
     State('interval-input', 'value'),
     State('appointment-type-dropdown', 'value'),
     State('plus-minus-days-input', 'value')]
)
def handle_calculate_button(n_clicks, name, start_date, interval_type, interval_value, appointment_type,
                            plus_minus_days):
    if n_clicks > 0 and name and start_date and interval_type and interval_value:
        plus_minus_days = int(plus_minus_days) if plus_minus_days else 0
        appointment_dates = calculate_appointment_dates(start_date, interval_type, interval_value, plus_minus_days)
        generate_pdf_report(name, appointment_dates, appointment_type, plus_minus_days)

        table_data = [['Cycle Number', 'Date', 'Plus Day Date', 'Minus Day Date']]
        for i in range(0, len(appointment_dates), 3):
            table_data.append([str(i // 3 + 1), appointment_dates[i], appointment_dates[i + 1], appointment_dates[i + 2]])

        table_rows = [html.Tr([html.Td(col) for col in row]) for row in table_data]
        table = html.Table(table_rows, className='table table-striped')

        return table, True

    return None, False


if _name_ == '_main_':
    app.run_server()