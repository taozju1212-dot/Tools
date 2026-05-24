# Motion Calculator

步进电机运动计算桌面工具。

## 功能

- 同步带直线轴、丝杆直线轴、自定义 `microsteps/mm` 换算。
- 项目首页：新建项目、打开项目、保存项目。
- 多运动轴管理：步进直线、步进旋转、步进丝杆。
- 多动作管理：每个动作绑定一个运动轴。
- 动作参数支持寄存器十进制值和物理单位双向换算。
- 每个动作支持多个距离，距离支持 `X_TARGET` 和物理单位双向换算。
- 自动判断 T 型 / 三角速度曲线，并显示多段时间点。
- 目标时间反算推荐 `Vmax`、`Acc`、`Dcc`。
- 保存 / 打开本地项目 JSON。
- `T 型三段反推` 暂不开发。
- `六点斜坡反算（待开发）` 入口已预留并灰显。

## 运行源码

```powershell
pip install -r requirements.txt
python main.py
```

## 打包

```powershell
pyinstaller --noconfirm --windowed --name MotionCalculator main.py
```

打包结果：

```text
dist/MotionCalculator/MotionCalculator.exe
```
