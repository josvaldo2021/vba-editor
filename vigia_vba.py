# -*- coding: utf-8 -*-
"""Monitor de erros do VBA: registra as caixas "Microsoft Visual Basic" de QUALQUER Excel aberto.

Cobre o que o `verificar` e os gates nao veem: o erro que aparece no Excel do usuario.
So OBSERVA: nunca clica, fecha nem derruba nada, e nao usa COM (o Excel fica bloqueado enquanto a
caixa esta aberta). Por isso registra o texto do erro, mas nao o modulo e a linha.

uso: python vigia_vba.py [--segundos N]     (sem --segundos, roda ate Ctrl+C)
log: %LOCALAPPDATA%\\pcp_gates\\erros_vba.log
"""
import datetime, os, subprocess, sys, time

import win32gui, win32process

LOG = os.path.join(os.environ["LOCALAPPDATA"], "pcp_gates", "erros_vba.log")


def texto_do_dialogo(h):
    partes = []
    win32gui.EnumChildWindows(h, lambda c, _: partes.append(win32gui.GetWindowText(c)), None)
    return " | ".join(" ".join(p.split()) for p in partes if p.strip())


def registrar(origem, texto):
    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    linha = "%s [%s] %s" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), origem, texto)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(linha + "\n")
    print(linha, flush=True)


def pids_excel():
    saida = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                           capture_output=True, text=True).stdout
    return {int(l.split('","')[1]) for l in saida.splitlines() if l.startswith('"EXCEL')}


def dialogos_vba(pids):
    achados = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "#32770":
            if win32process.GetWindowThreadProcessId(h)[1] in pids:
                if win32gui.GetWindowText(h).startswith("Microsoft Visual Basic"):
                    achados.append(h)

    win32gui.EnumWindows(cb, None)
    return achados


def main():
    limite = None
    if "--segundos" in sys.argv:
        limite = time.time() + float(sys.argv[sys.argv.index("--segundos") + 1])
    print("vigiando; log em %s" % LOG, flush=True)
    vistos = set()
    pids, proximo = set(), 0
    try:
        while limite is None or time.time() < limite:
            if time.time() >= proximo:
                pids, proximo = pids_excel(), time.time() + 2
            try:
                abertos = set(dialogos_vba(pids))
                for h in abertos - vistos:
                    registrar("monitor", texto_do_dialogo(h))
                vistos = abertos
            except Exception:
                pass  # a janela pode sumir no meio da leitura
            time.sleep(0.1)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
