# seewo-yuanshen

利用 vbs 劫持快捷方式来提前替换图片

### 需要更改的配置
- `./yuanshen.vbs` `swenlauncher`变量：希沃白板程序运行路径
- `./创建伪快捷方式.vbs` 第21行，调用`CreateShortcutOnDesktop`的第三个参数：希沃白板程序运行路径
- `./img.png` 需要替换的图片，默认原神启动

### 优点
- 不需要管理员权限（设置自启）
- 无需后台常驻

### 使用
- 将本文件夹放在一个**固定目录**后执行`创建伪快捷方式.vbs`
- 快捷方式生成后，不得删除`yuanshen.vbs`和`./img.png`
- 若要删除，删除伪快捷方式即可