"""
抽签小程序打包脚本
使用 PyInstaller 将程序打包成独立的可执行文件

✅ 支持系统：
- Windows: 生成 .exe 文件
- macOS: 生成可执行文件
- Linux: 生成可执行文件

⚠️  注意：
- 不同系统上打包生成的文件互不通用
- Windows系统打包的文件只能在Windows上运行
- macOS系统打包的文件只能在macOS上运行
- Linux系统打包的文件只能在Linux上运行
"""

import os
import sys
import subprocess
import platform

def check_dependencies():
    """检查必要的依赖"""
    print("🔍 检查依赖...")

    required_packages = [
        'pyinstaller',
        'pandas',
        'openpyxl',
        'PyQt6'
    ]

    missing_packages = []

    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} 已安装")
        except ImportError:
            print(f"❌ {package} 未安装")
            missing_packages.append(package)

    if missing_packages:
        print(f"\n⚠️  缺少以下依赖: {', '.join(missing_packages)}")
        print("正在安装依赖...")
        for package in missing_packages:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        print("✅ 依赖安装完成")

    return True

def build_exe():
    """使用 PyInstaller 打包"""
    print("\n📦 开始打包...")

    # 检测操作系统
    system = platform.system()
    print(f"🔍 检测到操作系统: {system}")

    # 显示打包目标
    if system == 'Windows':
        print(f"🎯 打包目标: Windows (.exe)")
    elif system == 'Darwin':
        print(f"🎯 打包目标: macOS")
    elif system == 'Linux':
        print(f"🎯 打包目标: Linux")
    else:
        print(f"🎯 打包目标: {system}")

    # PyInstaller 命令参数
    pyinstaller_cmd = [
        'pyinstaller',
        '--onefile',  # 打包成单个文件
        '--name=抽签小程序',  # 文件名
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不询问确认
    ]

    # 根据系统添加特定参数
    if system == 'Windows':
        pyinstaller_cmd.append('--windowed')  # 不显示控制台窗口
        pyinstaller_cmd.append('--icon=NONE')  # 图标
    elif system == 'Darwin':
        # macOS 特定参数
        pyinstaller_cmd.append('--windowed')  # 不显示终端窗口
        pyinstaller_cmd.append('--osx-bundle-identifier=com.抽签小程序')
        print("ℹ️  使用 macOS 打包参数")
    elif system == 'Linux':
        # Linux 也可以使用 --windowed
        pyinstaller_cmd.append('--windowed')

    pyinstaller_cmd.append('抽签小程序.py')

    try:
        subprocess.check_call(pyinstaller_cmd)
        print("\n✅ 打包成功！")

        # 根据系统显示不同的文件名
        if system == 'Windows':
            print(f"📁 可执行文件位置: dist/抽签小程序.exe")
        elif system == 'Darwin':
            print(f"📁 可执行文件位置: dist/抽签小程序")
            print(f"💡 提示: macOS上可以直接双击运行")
        elif system == 'Linux':
            print(f"📁 可执行文件位置: dist/抽签小程序")
            print(f"💡 提示: 运行命令: ./dist/抽签小程序")
        else:
            print(f"📁 可执行文件位置: dist/抽签小程序")

        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 打包失败: {e}")
        return False

def create_portable_package():
    """创建便携版包"""
    print("\n📦 创建便携版包...")

    # 创建目录结构
    package_dir = "抽签小程序_便携版"
    if os.path.exists(package_dir):
        import shutil
        shutil.rmtree(package_dir)

    os.makedirs(package_dir, exist_ok=True)

    # 复制可执行文件
    import shutil
    system = platform.system()

    # 根据系统确定文件名
    if system == 'Windows':
        exe_name = "抽签小程序.exe"
    elif system == 'Darwin':
        exe_name = "抽签小程序"
    else:
        exe_name = "抽签小程序"

    if os.path.exists(f"dist/{exe_name}"):
        shutil.copy(f"dist/{exe_name}", f"{package_dir}/{exe_name}")
        print(f"✅ 已复制可执行文件: {exe_name}")

        # 在macOS上，需要确保文件有执行权限
        if system == 'Darwin':
            try:
                os.chmod(f"{package_dir}/{exe_name}", 0o755)
                print("✅ 已设置执行权限")
            except Exception as e:
                print(f"⚠️  设置权限失败: {e}")

    # 复制 Excel 模板文件（如果存在）
    if os.path.exists("工作簿1.xlsx"):
        shutil.copy("工作簿1.xlsx", f"{package_dir}/工作簿1.xlsx")
        print("✅ 已复制 Excel 模板文件")

    # 复制使用说明
    if os.path.exists("使用说明.md"):
        shutil.copy("使用说明.md", f"{package_dir}/使用说明.md")
        print("✅ 已复制使用说明")

    # 复制快速入门
    if os.path.exists("快速入门.md"):
        shutil.copy("快速入门.md", f"{package_dir}/快速入门.md")
        print("✅ 已复制快速入门")

    print(f"\n✅ 便携版包已创建: {package_dir}/")

    # 根据系统显示提示
    if system == 'Windows':
        print(f"💡 Windows用户可以直接运行exe文件")
    elif system == 'Darwin':
        print(f"💡 macOS用户可以双击运行")
        print(f"   或在终端运行: cd 抽签小程序_便携版 && ./抽签小程序")
    elif system == 'Linux':
        print(f"💡 Linux用户在终端运行:")
        print(f"   cd 抽签小程序_便携版")
        print(f"   chmod +x 抽签小程序")
        print(f"   ./抽签小程序")
    else:
        print(f"⚠️  注意: 便携版中的文件是 {system} 系统可执行文件")

def main():
    """主函数"""
    print("=" * 60)
    print("   抽签小程序打包工具")
    print("   支持 Windows / macOS / Linux")
    print("=" * 60)

    # 显示系统信息
    system = platform.system()
    print(f"\n💻 当前系统: {system}")

    # 根据系统显示提示
    if system == 'Windows':
        print(f"🎯 将生成: Windows .exe 文件")
    elif system == 'Darwin':
        print(f"🎯 将生成: macOS 可执行文件")
    elif system == 'Linux':
        print(f"🎯 将生成: Linux 可执行文件")
    else:
        print(f"🎯 将生成: {system} 可执行文件")

    print(f"\n⚠️  注意: 生成的文件只能在 {system} 系统上运行")

    # 检查当前目录
    if not os.path.exists("抽签小程序.py"):
        print("\n❌ 错误: 未找到 '抽签小程序.py' 文件")
        print("请在包含抽签小程序的目录中运行此脚本")
        return

    # 检查依赖
    if not check_dependencies():
        print("❌ 依赖检查失败")
        return

    # 打包
    if not build_exe():
        print("❌ 打包失败")
        return

    # 创建便携版包
    create_portable_package()

    print("\n" + "=" * 60)
    print("✅ 打包完成！")
    print("=" * 60)

    # 根据系统显示不同的文件名和提示
    if system == 'Windows':
        print(f"📁 可执行文件: dist/抽签小程序.exe")
        print(f"📁 便携版包: 抽签小程序_便携版/")
        print(f"\n💡 Windows用户提示:")
        print(f"   - 单个exe文件可以直接运行")
        print(f"   - 便携版包包含使用说明，推荐使用")
    elif system == 'Darwin':
        print(f"📁 可执行文件: dist/抽签小程序")
        print(f"📁 便携版包: 抽签小程序_便携版/")
        print(f"\n💡 macOS用户提示:")
        print(f"   - 双击可执行文件即可运行")
        print(f"   - 首次运行可能需要在系统偏好设置中允许")
        print(f"   - 便携版包包含使用说明，推荐使用")
    elif system == 'Linux':
        print(f"📁 可执行文件: dist/抽签小程序")
        print(f"📁 便携版包: 抽签小程序_便携版/")
        print(f"\n💡 Linux用户提示:")
        print(f"   - 运行: ./dist/抽签小程序")
        print(f"   - 如果无法运行，执行: chmod +x ./dist/抽签小程序")
    else:
        print(f"📁 可执行文件: dist/抽签小程序")
        print(f"📁 便携版包: 抽签小程序_便携版/")

    print("=" * 60)

if __name__ == "__main__":
    main()
