import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
import base64
import io
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from dash.dependencies import Output, Input, State
import pandas as pd

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
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
        html.A(id='download-link', children='Download Excel', download='', href='', className='btn btn-primary')
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


def generate_excel_report(table_data):
    df = pd.DataFrame(table_data[1:], columns=table_data[0])
    excel_file = io.BytesIO()
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Appointment Report', index=False)
    excel_file.seek(0)
    excel_file_binary = excel_file.getvalue()
    encoded_excel = base64.b64encode(excel_file_binary).decode('utf-8')
    return f'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{encoded_excel}'


@app.callback(
    Output('appointment-table-div', 'children'),
    Output('alert', 'is_open'),
    Output('download-link', 'download'),
    Output('download-link', 'href'),
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
        table_data = [['Cycle Number', 'Date', 'Plus Day Date', 'Minus Day Date']]
        for i in range(0, len(appointment_dates), 3):
            table_data.append([str(i // 3 + 1), appointment_dates[i], appointment_dates[i + 1], appointment_dates[i + 2]])

        excel_link = generate_excel_report(table_data)

        table_rows = [html.Tr([html.Td(col) for col in row]) for row in table_data]
        table = html.Table(table_rows, className='table table-striped')

        return table, True, 'appointment_report.xlsx', excel_link

    return None, False, '', ''


if __name__ == '__main__':
    app.run_server(debug=True)
