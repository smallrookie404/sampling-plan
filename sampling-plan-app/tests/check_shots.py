from PIL import Image
import os

base = r"G:\codex工作台\采样计划\sampling-plan-app\tests\shots"
for name in ["1-main.png", "2-hazard.png", "3-items.png", "4-main-after-export.png"]:
    p = os.path.join(base, name)
    if not os.path.exists(p):
        print(name, "missing")
        continue
    im = Image.open(p).convert("RGB")
    w, h = im.size
    px = im.load()
    nonwhite = 0
    step = 7
    cnt = 0
    for y in range(0, h, step):
        for x in range(0, w, step):
            cnt += 1
            r, g, b = px[x, y]
            if not (r > 240 and g > 240 and b > 240):
                nonwhite += 1
    print(name, "size=%dx%d nonwhite%%=%.1f" % (w, h, 100 * nonwhite / cnt))
