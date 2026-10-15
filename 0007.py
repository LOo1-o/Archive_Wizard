# -*- coding: utf-8 -*-
"""
Сверка fc-файлов ОБДХ с планом выборки и между кварталами.

Зачем: протокол контроля (WKF_*.txt) сравнивает файлы ТОЛЬКО ДРУГ С ДРУГОМ.
Если участок пропал сразу из всех файлов, протокол этого не заметит — в нём
будет тишина. Строка "по плану= 0" означает, что план не подгрузился и
сверка с ним не выполнялась вовсе.

Поэтому здесь две проверки:
  1. факт против плана  — ловит участки, потерянные целиком;
  2. квартал против квартала — показывает, что изменилось.

Как читать номера домохозяйств (BUD):
  до 8000     — основная выборка, столько, сколько в плане;
  8000 и выше — дополнительные (замены), в контрольный счёт не входят.

Файлы плана лежат рядом с dbf:
  <тер>nsu.txt   — участки и районы
  <тер>bud.txt   — домохозяйства выборки
  <тер>okato.txt — населённые пункты
Кодировка у всех cp866, разделитель "#".

Запуск:
    python sravnit_fc.py fc30002641.dbf fc30002642.dbf
    python sravnit_fc.py fc30002642.dbf --plan .
"""
import collections
import os
import struct
import sys

MAIN_SAMPLE_LIMIT = 8000     # BUD >= этого числа — дополнительное домохозяйство
ENC = "cp866"


# ---------------------------------------------------------------- чтение DBF
def read_dbf(path):
    """Минимальный читатель DBF. Помеченные на удаление записи пропускаются."""
    data = open(path, "rb").read()
    ver, yy, mm, dd, nrec, hlen, rlen = struct.unpack("<BBBBIHH", data[:12])

    fields, pos = [], 32
    while data[pos:pos + 1] != b"\x0d":
        fd = data[pos:pos + 32]
        fields.append((fd[:11].split(b"\x00")[0].decode(ENC), fd[16]))
        pos += 32

    offsets, o = {}, 1
    for name, length in fields:
        offsets[name] = (o, length)
        o += length

    recs = []
    for i in range(nrec):
        raw = data[hlen + i * rlen: hlen + (i + 1) * rlen]
        if raw[0:1] == b"*":
            continue
        rec = {}
        for name, length in fields:
            s, ln = offsets[name]
            rec[name] = raw[s:s + ln].decode(ENC, errors="replace").strip()
        recs.append(rec)
    return recs


def read_lines(path):
    text = open(path, "rb").read().decode(ENC, errors="replace")
    return [l.split("#") for l in text.replace("\r\n", "\n").split("\n") if l.strip()]


# ---------------------------------------------------------------- план
def load_plan(folder):
    """
    Возвращает (план по участкам, названия участков, названия нас. пунктов).
    Ищет в папке файлы *nsu.txt, *bud.txt, *okato.txt — префикс любой.
    """
    def find(suffix):
        for name in sorted(os.listdir(folder)):
            if name.lower().endswith(suffix):
                return os.path.join(folder, name)
        return None

    su_names, plan, okato = {}, collections.Counter(), {}

    path = find("nsu.txt")
    if path:
        for row in read_lines(path):
            su_names[row[0].strip()] = row[1].split("-", 1)[-1].strip()

    path = find("okato.txt")
    if path:
        for row in read_lines(path):
            okato[row[0].strip()] = row[1].strip()

    # число домохозяйств на участок берём из ko-файла: там есть и участок,
    # и номер домохозяйства. bud.txt участок не содержит.
    path = find("ko_1.txt") or find("ko_2.txt") or find("ko_3.txt")
    if path:
        seen = set()
        for row in read_lines(path):
            # в ko_1 участок в 7-м поле, в ko_2/ko_3 — в 6-м
            idx = 6 if len(row) == 18 else 5
            key = (row[idx], row[idx + 1])
            if key not in seen:
                seen.add(key)
                plan[row[idx]] += 1

    return plan, su_names, okato


# ---------------------------------------------------------------- факт
def fact_by_su(recs):
    """{участок: (основных, дополнительных)}"""
    main, extra = collections.Counter(), collections.Counter()
    for r in recs:
        su = r["SU"].lstrip("0")
        if int(r["BUD"]) < MAIN_SAMPLE_LIMIT:
            main[su] += 1
        else:
            extra[su] += 1
    return main, extra


# ---------------------------------------------------------------- отчёты
def report_plan(path_dbf, folder):
    plan, su_names, okato = load_plan(folder)
    if not plan:
        print("В папке %s не найдены файлы плана (nsu/ko). Сверка с планом "
              "пропущена." % folder)
        return
    recs = read_dbf(path_dbf)
    main, extra = fact_by_su(recs)

    print("=" * 76)
    print("СВЕРКА С ПЛАНОМ ВЫБОРКИ: %s" % os.path.basename(path_dbf))
    print("=" * 76)
    print("%-8s %-40s %6s %6s %s" % ("участок", "район", "план", "факт", ""))
    print("-" * 76)

    tp = tf = 0
    lost, short = [], []
    for su in sorted(set(plan) | set(main)):
        p, f = plan[su], main[su]
        tp += p
        tf += f
        mark = ""
        if p and not f:
            mark = "  <-- ПРОПАЛ ЦЕЛИКОМ"
            lost.append(su)
        elif f < p:
            mark = "  <-- недобор %d" % (p - f)
            short.append(su)
        elif f > p:
            mark = "  <-- сверх плана %d" % (f - p)
        if extra[su]:
            mark += "  (+%d доп.)" % extra[su]
        print("%-8s %-40s %6d %6d%s" % (su, su_names.get(su, "?")[:40], p, f, mark))

    print("-" * 76)
    print("%-49s %6d %6d" % ("ИТОГО основных", tp, tf))
    if tf != tp:
        print()
        print("НЕ ХВАТАЕТ %d домохозяйств из %d." % (tp - tf, tp))
    if lost:
        print("Участки, пропавшие целиком: %s" % ", ".join(lost))
    if short:
        print("Участки с недобором: %s" % ", ".join(short))
    print()
    print("Внимание: протокол контроля сверяет файлы только друг с другом.")
    print("Участок, пропавший сразу из всех файлов, в протокол НЕ попадёт.")


def report_diff(path_old, path_new, folder=None):
    su_names = load_plan(folder)[1] if folder else {}
    old, new = read_dbf(path_old), read_dbf(path_new)
    mo, eo = fact_by_su(old)
    mn, en = fact_by_su(new)

    print("=" * 76)
    print("%-30s записей %4d, участков %2d" % (os.path.basename(path_old),
                                               len(old), len(mo)))
    print("%-30s записей %4d, участков %2d" % (os.path.basename(path_new),
                                               len(new), len(mn)))
    print("=" * 76)
    print("%-8s %-34s %-13s %-13s" % ("участок", "район", "было", "стало"))
    print("-" * 76)

    def show(m, e, su):
        if not m[su] and not e[su]:
            return "нет"
        s = "%d осн." % m[su]
        if e[su]:
            s += " +%d" % e[su]
        return s

    for su in sorted(set(mo) | set(mn) | set(eo) | set(en)):
        mark = ""
        if mo[su] and not mn[su]:
            mark = "  <-- ПРОПАЛ"
        elif mn[su] and not mo[su]:
            mark = "  <-- новый"
        elif mo[su] != mn[su]:
            mark = "  <-- основная выборка изменилась"
        print("%-8s %-34s %-13s %-13s%s" % (su, su_names.get(su, "")[:34],
                                            show(mo, eo, su),
                                            show(mn, en, su), mark))
    print("-" * 76)
    print("основных: было %d, стало %d" % (sum(mo.values()), sum(mn.values())))


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    folder = "."
    if "--plan" in sys.argv:
        i = sys.argv.index("--plan")
        if i + 1 < len(sys.argv):
            folder = sys.argv[i + 1]

    if len(args) >= 2:
        report_diff(args[0], args[1], folder)
        print()
        report_plan(args[1], folder)
    elif len(args) == 1:
        report_plan(args[0], folder)
    else:
        print(__doc__)