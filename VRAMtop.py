# -*- coding: utf-8 -*-
from getVRAMInfo import get_vram_info, get_vram_usages


if __name__ == '__main__':
    import time

    print(get_vram_info())

    # targets = ['Dedicated', 'Local']
    targets = ['Local']
    for target in targets:
        begin = time.perf_counter()
        usages = get_vram_usages(target=target)
        print(f'{target} elapsed: {time.perf_counter() - begin:.3f}s')

        for name in sorted(usages, key=lambda n: usages[n], reverse=True):
            mem = usages[name]
            if mem != 0:
                print(f'{name:24}\t{mem / 1024 / 1024:8.1f} MB')
