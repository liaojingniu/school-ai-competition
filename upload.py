import serial
import time

# ESP32-C3 MicroPython 控制逻辑 (专用于双路继电器：K1风扇 + K2 LED灯)
# 原生支持按键: 1(风扇开), 2(风扇关), 3(LED开), 4(LED关), 5(全开), 0(全关), P(极性), T(全引脚雷达)
ESP_CODE = '''import sys, select, time
from machine import Pin

# 风扇主控 GPIO 4 (冗余 GPIO 2)
# LED 主控 GPIO 5 (冗余 GPIO 3)
fan_pins = [Pin(4, Pin.OUT), Pin(2, Pin.OUT)]
led_pins = [Pin(5, Pin.OUT), Pin(3, Pin.OUT)]

# 极性状态: 0=常规低电平吸合(默认), 1=反向高电平吸合
polarity_mode = 0

fan_on = False
led_on = False

def apply_device_states():
    v_fan = 0 if fan_on else 1
    v_led = 0 if led_on else 1
    if polarity_mode == 1:
        v_fan = 1 - v_fan
        v_led = 1 - v_led
    for p in fan_pins:
        p.value(v_fan)
    for p in led_pins:
        p.value(v_led)

apply_device_states()

poll_obj = select.poll()
poll_obj.register(sys.stdin, select.POLLIN)

print("ESP32_READY")

while True:
    if poll_obj.poll(10):
        ch = sys.stdin.read(1)
        if ch in ('1', 'A'):
            fan_on = True
            apply_device_states()
        elif ch in ('2', 'a'):
            fan_on = False
            apply_device_states()
        elif ch in ('3', 'B'):
            led_on = True
            apply_device_states()
        elif ch in ('4', 'b'):
            led_on = False
            apply_device_states()
        elif ch == '5':
            fan_on = True
            led_on = True
            apply_device_states()
        elif ch in ('0', 'X'):
            fan_on = False
            led_on = False
            apply_device_states()
        elif ch == 'P':
            polarity_mode = 1 if polarity_mode == 0 else 0
            apply_device_states()
        elif ch == 'S':
            sys.stdout.write("STAT:{},{},{}\\n".format(
                1 if fan_on else 0,
                1 if led_on else 0,
                polarity_mode
            ))
        elif ch == 'T':
            # 全引脚扫描雷达模式：对 GPIO 0~10 轮流触发 0.8 秒，帮助听继电器咔哒声
            test_list = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
            for p_num in test_list:
                sys.stdout.write("SCAN_PIN:{}\\n".format(p_num))
                try:
                    tp = Pin(p_num, Pin.OUT)
                    tp.value(0)
                    time.sleep(0.8)
                    tp.value(1)
                    time.sleep(0.2)
                except:
                    pass
            sys.stdout.write("SCAN_DONE\\n")
            apply_device_states()
'''

PORT = 'COM3'
BAUD = 115200

def flash_esp():
    print(f"正在连接 {PORT}...")
    ser = serial.Serial(PORT, BAUD, timeout=1)
    time.sleep(0.5)

    # 1. 打断并进入 raw REPL (Ctrl+C 然后 Ctrl+A)
    ser.write(b'\x03\x03')
    time.sleep(0.3)
    ser.write(b'\x01')
    time.sleep(0.3)

    raw_resp = ser.read_all()
    if b'raw REPL' not in raw_resp:
        ser.write(b'\x03\x01')
        time.sleep(0.3)

    print("已进入 MicroPython 原始交互模式，正在写入 main.py...")

    # 构建写入文件的代码
    write_script = "content = " + repr(ESP_CODE) + "\n" + "with open('main.py', 'w') as f:\n    f.write(content)\nprint('FILE_WRITTEN_SUCCESSFULLY')\n"

    ser.write(write_script.encode('utf-8'))
    ser.write(b'\x04')
    time.sleep(1.0)

    result = ser.read_all().decode('utf-8', errors='replace')
    print("写入反馈:", result)

    # 发送 Ctrl+D 触发软复位执行新的 main.py
    ser.write(b'\x02') # 退出 raw REPL
    time.sleep(0.2)
    ser.write(b'\x04') # 软复位
    time.sleep(0.8)

    ser.close()
    if 'FILE_WRITTEN_SUCCESSFULLY' in result:
        print("\n[成功] 纯按键双路高容错版 main.py 已成功写入 ESP32-C3 并自启运行！")
    else:
        print("\n[完成] 写入指令已发送，请检验硬件。")

if __name__ == '__main__':
    flash_esp()