import sys
import os
import openpyxl


def main():
    if len(sys.argv) < 2 or not os.path.exists(sys.argv[1]):
        print("输入文件错误", file=sys.stderr)
        sys.exit()
    filename = sys.argv[1]
    wb = openpyxl.load_workbook(filename)
    stats = {}
    team_count = 0
    for ws in wb:
        for column in ws.columns:
            i = 0
            l = len(column) - 6
            while i < l:
                o = True
                for j in range(6):
                    if not (isinstance(column[i+j].value, str) and column[i+j].value.isalpha()):
                        o = False
                        break
                if o:
                    team_count += 1
                    for j in range(6):
                        v = column[i+j].value
                        if v not in stats:
                            stats[v] = [0, 0]
                        stats[v][0] += 1
                        if column[i+j].style.font.color.rgb=='FFFF0000':
                            stats[v][1] += 1
                    i += 5
                i += 1
    total = team_count / 100
    results = []
    for each in stats:
        results.append((
            each,
            stats[each][0],
            stats[each][1],
            stats[each][0] / total,
            stats[each][1] / total
        ))
    results.sort(key=lambda a: a[1], reverse=True)
    with open('result.txt', 'w', encoding='utf-8') as fo:
        for each in results:
            print('%s\t%d\t%d\t%.2f%%\t%.2f%%' % each, file=fo)


if __name__ == '__main__':
    main()
