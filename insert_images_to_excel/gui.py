import tkinter
import tkinter.simpledialog


def create_simple_dialog(title: str, prompt: str) -> str:

    ROOT = tkinter.Tk()
    ROOT.withdraw()

    user_input = tkinter.simpledialog.askstring(
        title=title,
        prompt=prompt)

    return user_input
