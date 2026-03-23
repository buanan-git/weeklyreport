# cleanup.py
import os
import shutil
from pathlib import Path

def cleanup():
    """清理混乱的目录结构"""
    print("="*60)
    print("清理打包文件 - 开始")
    print("="*60)
    
    root_dir = Path.cwd()
    scripts_dir = root_dir / "scripts"
    
    # 1. 删除scripts下的dist目录（应该移动到根目录）
    scripts_dist = scripts_dir / "dist"
    if scripts_dist.exists():
        print(f"找到 scripts/dist/ 目录，正在移动到根目录...")
        
        # 如果根目录已有dist，先备份
        root_dist = root_dir / "dist"
        if root_dist.exists():
            backup_dir = root_dir / "dist_backup"
            if backup_dir.exists():
                shutil.rmtree(backup_dir)
            root_dist.rename(backup_dir)
            print(f"根目录原有dist已备份到: dist_backup/")
        
        # 移动scripts/dist到根目录
        scripts_dist.rename(root_dist)
        print(f"✅ 已移动: scripts/dist/ → dist/")
    
    # 2. 删除根目录下多余的dist文件（如果还有）
    root_dist_exe = root_dir / "周报助手.exe"
    if root_dist_exe.exists():
        root_dist_exe.unlink()
        print(f"✅ 删除根目录下的多余exe文件")
    
    # 3. 删除build目录
    build_dir = root_dir / "build"
    if build_dir.exists():
        shutil.rmtree(build_dir)
        print(f"✅ 删除 build/ 目录")
    
    # 4. 删除__pycache__目录
    for pycache in root_dir.rglob("__pycache__"):
        if pycache.is_dir():
            shutil.rmtree(pycache)
            print(f"✅ 删除: {pycache}")
    
    # 5. 删除.pyc文件
    for pyc in root_dir.rglob("*.pyc"):
        pyc.unlink()
        print(f"✅ 删除: {pyc}")
    
    # 6. 检查最终结构
    print("\n" + "="*60)
    print("清理完成！当前目录结构：")
    print("="*60)
    
    # 列出重要目录
    important_dirs = ['scripts', 'config', 'dist']
    for dir_name in important_dirs:
        dir_path = root_dir / dir_name
        if dir_path.exists():
            size = sum(f.stat().st_size for f in dir_path.rglob('*') if f.is_file()) / 1024
            print(f"📁 {dir_name}/ ({size:.1f} KB)")
            
            # 列出关键文件
            if dir_name == 'dist':
                files = list(dir_path.glob('*'))
                for f in files[:5]:  # 只显示前5个
                    if f.is_file():
                        print(f"   ├─ {f.name}")
    
    print("\n✅ 清理完成！")
    print("\n下一步：")
    print("1. 检查 dist/ 目录下的文件是否完整")
    print("2. 运行 dist/调试.bat 测试程序")
    print("3. 如果一切正常，可以删除 dist_backup/ 目录")

if __name__ == "__main__":
    response = input("此操作将清理打包文件，确定继续？(y/n): ")
    if response.lower() == 'y':
        cleanup()
    else:
        print("已取消")