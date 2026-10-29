import psutil
import openpyxl
import time
import matplotlib.pyplot as plt

# Function to collect system data
def collect_system_data(): #indentetion
    cpu_usage = psutil.cpu_percent(interval=1)
    ram_usage = psutil.virtual_memory().percent
    disk_usage = psutil.disk_usage('/').percent
    battery = psutil.sensors_battery()
    power_plugged = battery.power_plugged if battery else False
    network_speed = psutil.net_io_counters().bytes_sent + psutil.net_io_counters().bytes_recv

    return cpu_usage, ram_usage, disk_usage, network_speed

# ... (create Excel workbook and add worksheet)

cpu_values = []  # List to store CPU usage values
timestamps = []  # List to store timestamps

try:
    for _ in range(60):
        timestamp = time.strftime("%I:%M:%S %p")  # Time for x-axis
        cpu, _, _, _ = collect_system_data()

        cpu_values.append(cpu)
        timestamps.append(timestamp)

        # Plot real-time CPU usage
        plt.clf()
        plt.plot(timestamps, cpu_values, marker='o')
        plt.xticks(rotation=45)
        plt.xlabel("Time")
        plt.ylabel("CPU Usage (%)")
        plt.title("Real-Time CPU Usage")
        plt.tight_layout()
        plt.savefig("cpu_usage_plot.png")  # Save plot as an image
        plt.pause(1)  # Pause for 1 second

except KeyboardInterrupt:
    print("Monitoring stopped by user.")
