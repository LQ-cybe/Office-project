"""将 AppIcon.ico 从黑色重绘为亮蓝 #3370FF（与视图切换按钮同色），保留多尺寸。"""
from PIL import Image
import os, sys

SRC = r"D:\编程开发\开源项目\多维表分析软件\MultiTableAddin\src\MultiTableAddin\Resources\Icons\AppIcon.ico"
DST = r"D:\编程开发\开源项目\多维表分析软件\MultiTableAddin\src\MultiTableAddin\Resources\Icons\AppIcon.ico"
DST2 = r"D:\编程开发\开源项目\多维表分析软件\图标汇总\功能区与窗口图标\AppIcon.ico"

# 蓝色：与 --accent #3370FF 一致（视图切换按钮文字色）
BLUE = (51, 112, 255)

# 打开原始 ICO，取最大尺寸作为高质量源；用 sizes 列表保留全部原尺寸
im = Image.open(SRC).convert("RGBA")
sizes = sorted(set(im.info.get("sizes", [(256, 256)])))
print("原始尺寸：", sizes)

# 1) 对最大尺寸（256）做精确重绘：任何非透明像素 → 纯蓝 #3370FF，保留 alpha
#    原图是纯黑 (R=G=B=0) + alpha 形状；重绘后 alpha 形状不变，
#    浅色高光（小圆点/反走样边缘）通过 alpha 与背景合成自然呈现。
src = im.copy()
px = src.load()
for y in range(src.size[1]):
    for x in range(src.size[0]):
        r, g, b, a = px[x, y]
        if a > 0:
            px[x, y] = (BLUE[0], BLUE[1], BLUE[2], a)

# 2) 保存为 ICO，列出全部 6 个尺寸（16/32/48/64/128/256），PIL 从 256 源 LANCZOS 缩放
out_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
src.save(DST, format="ICO", sizes=out_sizes)
print("已写入：", DST, "大小=", os.path.getsize(DST), "bytes")

# 同步到图标汇总目录（仅参考，不嵌入 DLL）
src.save(DST2, format="ICO", sizes=out_sizes)
print("已写入：", DST2, "大小=", os.path.getsize(DST2), "bytes")

# 导出 256 PNG 用于人眼校验
src.save("/tmp/appicon_256_blue.png", format="PNG")
print("导出预览：/tmp/appicon_256_blue.png")
