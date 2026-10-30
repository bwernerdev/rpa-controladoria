"""Lista os controles da janela do SRP. Somente leitura: nao digita nem clica."""
import sys

from pywinauto import Desktop


def main():
    filtro = (sys.argv[1] if len(sys.argv) > 1 else "SRP").lower()
    janelas = [w for w in Desktop(backend="uia").windows() if filtro in w.window_text().lower()]
    if not janelas:
        print(f"Nenhuma janela com '{filtro}' no titulo. Janelas abertas:")
        for w in Desktop(backend="uia").windows():
            if w.window_text():
                print(" -", w.window_text())
        return 1
    for janela in janelas:
        print(f"== {janela.window_text()} ({janela.element_info.control_type})")
        for c in janela.descendants():
            info = c.element_info
            if info.control_type in ("Edit", "ComboBox", "Button", "Text", "Spinner", "Document"):
                print(f"{info.control_type:9} nome={info.name!r} id={info.automation_id!r} classe={info.class_name!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
