import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objs as go
import dash_table
import pandas as pd
import math

# Step 1: Specify the external Excel file
excel_filename = 'pv_power_data.xlsx'

# Step 2: Read data from the external Excel file
df = pd.read_excel(excel_filename)

# Initialize the Dash app
app = dash.Dash(__name__)

# Initial values for sliders
initial_light_power = 1000  # Initial value for light power in lumens
initial_distance = 10  # Initial value for distance in meters
initial_cell_temp = 22  # Initial value for normal cell operating temperature in °C
initial_ambient_temp = 18  # Initial value for ambient temperature in °C
initial_cleanings = 5  # Initial value for number of cleanings
initial_solar_area = 5  # Initial value for area of solar panel in m²
initial_battery_capacity = 50  # Initial value for battery capacity in Ah
initial_light_power_watts = 50  # Initial value for light power in watts
initial_time = 0  # Initial value for time slider (6 PM)

# Define the time labels from 6 PM to 6 AM
time_labels = [
    '06:00 PM', '07:00 PM', '08:00 PM', '09:00 PM', '10:00 PM', '11:00 PM', 
    '12:00 AM', '01:00 AM', '02:00 AM', '03:00 AM', '04:00 AM', '05:00 AM', '06:00 AM'
]

# Define styles
styles = {
    'backgroundColor': '#f0f0f5',  # light purple background color
    'padding': '20px'
}

# Define the layout of the app
app.layout = html.Div(style=styles, children=[
    html.H1('Solar Street Light Simulation', style={'textAlign': 'center'}),

    # Lux calculation section
    html.H2('Lux Calculation'),
    html.Label('Light Power (lumens)'),
    dcc.Slider(id='light-power', min=0, max=2000, step=100, value=initial_light_power),
    html.Label('Distance from Light (meters)'),
    dcc.Slider(id='distance', min=1, max=20, step=1, value=initial_distance),
    html.Div(id='lux-output'),

    # Temperature and Soiling loss section
    html.H2('Temperature and Soiling Loss'),
    html.Label('Normal Cell Operating Temperature (°C)'),
    dcc.Slider(id='cell-temp', min=-10, max=40, step=2, value=initial_cell_temp),
    html.Label('Ambient Temperature (°C)'),
    dcc.Slider(id='ambient-temp', min=-10, max=40, step=2, value=initial_ambient_temp),
    html.Label('Number of Cleanings'),
    dcc.Slider(id='cleanings', min=0, max=10, step=1, value=initial_cleanings),

    # Area of Solar Panel slider
    html.Label('Area of Solar Panel (m²)'),
    dcc.Slider(id='solar-area', min=1, max=10, step=1, value=initial_solar_area),

    # Solar Panel Output graph
    html.H2('Solar Panel Output Over Time'),
    dcc.Graph(id='solar-panel-output-graph'),

    # Battery Charging Level graph
    html.H2('Battery Charging Level Over Time'),
    html.Label('Battery Capacity (Ah)'),
    dcc.Slider(id='battery-capacity', min=1, max=100, step=1, value=initial_battery_capacity),
    dcc.Graph(id='battery-charging-level-graph'),

    # Power Consumption Over Time section
    html.H2('Power Consumption Over Time'),
    html.Label('Light Power (watts)'),
    dcc.Slider(id='light-power-watts', min=0, max=100, step=1, value=initial_light_power_watts),
    html.Label('Select Brightness Model'),
    dcc.Dropdown(
        id='brightness-model',
        options=[
            {'label': 'Full Brightness', 'value': 'full'},
            {'label': 'Dimming Model - Slot 1', 'value': 'dim_slot1'},
            {'label': 'Dimming Model - Slot 2', 'value': 'dim_slot2'}
        ],
        value='full'
    ),
    dcc.Graph(id='power-consumption-graph'),

    # Battery Discharge Level section
    html.H2('Battery Discharge Level Over Time'),
    dcc.Graph(id='battery-discharge-level-graph'),

    # Timing slider for battery charge level
    html.H2('Select Time for Battery Charge Level'),
    dcc.Slider(
        id='time-slider',
        min=0,
        max=12,
        marks={i: label for i, label in enumerate(time_labels)},
        value=initial_time
    ),

    # Battery Gauge Chart
    html.H2('Battery Charge Level Gauge'),
    dcc.Graph(id='battery-charge-gauge'),

    # Data table section
    html.H2('Data Table'),
    dcc.RadioItems(
        id='view-selector',
        options=[
            {'label': 'Daily', 'value': 'day'},
            {'label': 'Monthly', 'value': 'month'},
            {'label': 'Hourly', 'value': 'hour'}
        ],
        value='month',
        labelStyle={'display': 'inline-block'}
    ),
    dcc.Dropdown(
        id='dropdown',
        placeholder="Select a filter",
    ),
    dash_table.DataTable(
        id='table',
        columns=[
            {"name": "Month", "id": "Month"},
            {"name": "Day", "id": "Day"},
            {"name": "Hour", "id": "Hour"},
            {"name": "Irradiance", "id": "Irradiance"},
            {"name": "PV_Power", "id": "PV_Power"}
        ],
        style_table={'overflowX': 'auto'},
        style_cell={
            'height': 'auto',
            'minWidth': '0px', 'maxWidth': '180px',
            'whiteSpace': 'normal'
        }
    ),
    html.Button(id='submit-button', n_clicks=0, children='Refresh Data')
])


# Callback to update Lux calculation output
@app.callback(
    Output('lux-output', 'children'),
    [Input('light-power', 'value'),
     Input('distance', 'value')]
)
def update_lux_output(light_power, distance):
    lux_per_m2 = light_power / (4 * math.pi * distance ** 2)
    return f'Lux/m²: {lux_per_m2:.2f}'


# Callback to update Solar Panel Output and Battery Charging Level graphs
@app.callback(
    [Output('solar-panel-output-graph', 'figure'),
     Output('battery-charging-level-graph', 'figure')],
    [Input('light-power', 'value'),
     Input('distance', 'value'),
     Input('cell-temp', 'value'),
     Input('ambient-temp', 'value'),
     Input('cleanings', 'value'),
     Input('solar-area', 'value'),
     Input('battery-capacity', 'value')]
)
def update_calculations(light_power, distance, cell_temp, ambient_temp, cleanings, solar_area, battery_capacity):
    # Calculate Lux/m^2
    lux_per_m2 = light_power / (4 * math.pi * distance ** 2)

    # Calculate Temperature Loss
    temp_loss = (ambient_temp - cell_temp) * 0.41

    # Calculate Soiling Loss
    initial_soiling_loss = 5  # Initial soiling loss percentage
    cleaning_factor = 0.9  # Cleaning factor
    soiling_loss = initial_soiling_loss * (1 - cleaning_factor) ** cleanings

    # Calculate Solar Panel Output
    irradiance = {
        '06:00 AM': 200, '07:00 AM': 400, '08:00 AM': 600, '09:00 AM': 800,
        '10:00 AM': 1000, '11:00 AM': 1100, '12:00 PM': 1200, '01:00 PM': 1100,
        '02:00 PM': 1000, '03:00 PM': 800, '04:00 PM': 600, '05:00 PM': 400, '06:00 PM': 200
    }
    times = list(irradiance.keys())
    solar_output = [irradiance[time] * (100 - temp_loss - soiling_loss) / 100 * solar_area for time in times]

    # Prepare data for the solar panel output graph
    solar_panel_output_fig = {
        'data': [{'x': times, 'y': solar_output, 'type': 'line', 'name': 'Solar Panel Output (W)'}],
        'layout': {'title': 'Solar Panel Output Over Time', 'xaxis': {'title': 'Time'}, 'yaxis': {'title': 'Power (W)'}}
    }

    # Calculate Battery Charging Level
    charging_current = [solar_power / 12 for solar_power in solar_output]
    time_hours = list(range(len(times)))
    Q_charging = [charging_current[i] * time_hours[i] for i in range(len(times))]
    Q_total = battery_capacity

    # Calculate Charging Level (%)
    charging_level = [(Q / Q_total) * 100 for Q in Q_charging]

    # Ensure charging level does not exceed 99% (stop condition)
    charging_level = [min(level, 99) for level in charging_level]

    # Prepare data for the battery charging level graph
    battery_charging_level_fig = {
        'data': [{'x': times, 'y': charging_level, 'type': 'line', 'name': 'Battery Charging Level (%)'}],
        'layout': {'title': 'Battery Charging Level Over Time', 'xaxis': {'title': 'Time'}, 'yaxis': {'title': 'Charging Level (%)'}}
    }

    return solar_panel_output_fig, battery_charging_level_fig


# Callback to update Power Consumption graph and calculate Battery Discharge and Charge Levels
@app.callback(
    [Output('power-consumption-graph', 'figure'),
     Output('battery-discharge-level-graph', 'figure'),
     Output('battery-charge-gauge', 'figure')],
    [Input('light-power-watts', 'value'),
     Input('brightness-model', 'value'),
     Input('battery-capacity', 'value'),
     Input('time-slider', 'value')]
)
def update_power_consumption_and_battery_levels(light_power_watts, brightness_model, battery_capacity, selected_time):
    if brightness_model == 'full':
        power_consumption = [light_power_watts] * 13
    elif brightness_model == 'dim_slot1':
        # Dimming Model - Slot 1
        power_consumption = []
        for i in range(13):
            if i <= 5:
                power_consumption.append(light_power_watts)  # Full brightness for first 5 hours
            else:
                power_consumption.append(light_power_watts * 0.5)  # Half brightness for next 7 hours
    elif brightness_model == 'dim_slot2':
        # Dimming Model - Slot 2
        power_consumption = []
        for i in range(13):
            if i <= 6:
                power_consumption.append(light_power_watts * 0.5)  # Half brightness for first 7 hours
            else:
                power_consumption.append(light_power_watts)  # Full brightness for next 5 hours

    # Calculate power consumption figure
    power_consumption_fig = {
        'data': [{'x': time_labels, 'y': power_consumption, 'type': 'line', 'name': 'Power Consumption (W)'}],
        'layout': {'title': 'Power Consumption Over Time', 'xaxis': {'title': 'Time'}, 'yaxis': {'title': 'Power Consumption (W)'}}
    }

    # Calculate Battery Discharge Level
    energy_consumption = [sum(power_consumption[:i+1]) for i in range(13)]
    total_battery_capacity = battery_capacity * 12  # Assuming 12V battery
    discharge_level = [(energy / total_battery_capacity) * 100 for energy in energy_consumption]

    # Invert discharge level (battery discharge increases over time)
    discharge_level = [100 - level for level in discharge_level]

    # Ensure discharge level does not exceed 99%
    discharge_level = [min(level, 99) for level in discharge_level]

    # Ensure discharge level does not drop below 1%
    discharge_level = [max(level, 1) for level in discharge_level]

    # Calculate Battery Charge Level
    charge_level = [100 - level for level in discharge_level]

    # Prepare data for the battery discharge level graph
    battery_discharge_level_fig = {
        'data': [{'x': time_labels, 'y': discharge_level, 'type': 'line', 'name': 'Battery Discharge Level (%)'}],
        'layout': {'title': 'Battery Discharge Level Over Time', 'xaxis': {'title': 'Time'}, 'yaxis': {'title': 'Discharge Level (%)'}}
    }

    # Create battery gauge chart for charge level (opposite of discharge level)
    battery_charge_gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=100-charge_level[selected_time],
        gauge={
            'axis': {'range': [0,100]},
            'bar': {'color': "darkblue"},
            'steps': [
                {'range': [0, 25], 'color': "red"},
                {'range': [25, 50], 'color': "orange"},
                {'range': [50, 75], 'color': "yellow"},
                {'range': [75, 100], 'color': "green"}
            ],
        }
    ))

    battery_charge_gauge.update_layout(
        title='Battery Charge Level',
        margin={'l': 20, 'r': 20, 't': 40, 'b': 20}
    )

    return power_consumption_fig, battery_discharge_level_fig, battery_charge_gauge


# Callback to update dropdown options based on view selector
@app.callback(
    Output('dropdown', 'options'),
    [Input('view-selector', 'value')]
)
def update_dropdown(view):
    if view == 'month':
        options = [{'label': f'Month {i}', 'value': i} for i in df['Month'].unique()]
    elif view == 'day':
        options = [{'label': f'Day {i}', 'value': i} for i in df['Day'].unique()]
    elif view == 'hour':
        options = [{'label': f'Hour {i}', 'value': i} for i in df['Hour'].unique()]
    return options


# Callback to update data table based on selected filter
@app.callback(
    Output('table', 'data'),
    [Input('submit-button', 'n_clicks'),
     Input('dropdown', 'value'),
     Input('view-selector', 'value')]
)
def refresh_table(n_clicks, filter_value, view):
    df = pd.read_excel(excel_filename)
    
    if view == 'month' and filter_value:
        filtered_df = df[df['Month'] == filter_value]
    elif view == 'day' and filter_value:
        filtered_df = df[df['Day'] == filter_value]
    elif view == 'hour' and filter_value:
        filtered_df = df[df['Hour'] == filter_value]
    else:
        filtered_df = df

    return filtered_df.to_dict('records')


# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
