"""Linha de comando do robo do SRP. Ex.: python src/preencher.py planilhas/arquivo.xlsx --dry-run"""
import argparse
import sys
import threading

from core import descrever, executar, ler_planilha, validar


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("planilha")
    ap.add_argument("--dry-run", action="store_true", help="mostra as teclas sem enviar nada")
    ap.add_argument("--janela", default="Desktop Viewer", help="trecho do titulo da janela do Citrix")
    ap.add_argument("--pausa", type=float, default=0.25, help="segundos entre teclas")
    ap.add_argument("--render", type=float, default=4.0, help="tempo maximo (s) esperando o foco mudar de campo")
    ap.add_argument("--espera", type=int, default=5, help="segundos para posicionar o cursor no Item")
    ap.add_argument("--inicio", type=int, default=2, help="primeira linha da planilha a processar")
    ap.add_argument("--limite", type=int, default=0, help="maximo de linhas (0 = todas)")
    args = ap.parse_args()
    try:
        registros = [r for r in ler_planilha(args.planilha) if r["linha"] >= args.inicio]
    except ValueError as erro:
        raise SystemExit(str(erro))
    if args.limite:
        registros = registros[: args.limite]
    if not registros:
        raise SystemExit("Nenhuma linha para processar.")
    erros = validar(registros)
    if erros:
        raise SystemExit("Planilha invalida, nada foi enviado:\n  " + "\n  ".join(erros))
    print(f"{len(registros)} linha(s) a processar.")
    if args.dry_run:
        for r in registros:
            print(f"linha {r['linha']}: {descrever(r)}")
        return 0
    return executar(registros, args.janela, args.pausa, args.render, args.espera, threading.Event(), print)


if __name__ == "__main__":
    sys.exit(main())
