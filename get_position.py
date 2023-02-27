import keyboard
import mouse

while True:
    keyboard.wait("spacebar")
    print(mouse.get_position())