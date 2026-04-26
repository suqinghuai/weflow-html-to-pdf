import configparser
import json
import os
import re
import sys
import threading
import time
from pathlib import Path

from playwright.sync_api import sync_playwright


def load_cut_size(config_path: Path, default_size: int = 4000) -> int:
    """从配置文件加载分割大小。"""
    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")
    try:
        value = config.get("BASE", "cutmessages")
        return max(1, int(value))
    except Exception:
        return default_size


def extract_data_block(html_text: str):
    """提取 HTML 中的 WEFLOW_DATA 数据块。"""
    pattern = re.compile(
        r"^(?P<indent>\s*)window\.WEFLOW_DATA\s*=\s*\[(?P<data>.*?)\];",
        re.DOTALL | re.MULTILINE,
    )
    match = pattern.search(html_text)
    if not match:
        return None, None, None
    return match.group("data"), match.group("indent"), match.span()


def parse_data_list(raw_data: str):
    """解析消息列表。"""
    return json.loads("[" + raw_data + "]")


def build_data_block(messages, indent: str) -> str:
    """构建替换后的数据块。"""
    lines = [json.dumps(item, ensure_ascii=False, separators=(",", ": ")) for item in messages]
    joined = ",\n".join(lines)
    return f"{indent}window.WEFLOW_DATA = [\n{joined}\n{indent}];"


def update_counts(html_text: str, count: int) -> str:
    """更新页面上的消息计数。"""
    html_text = re.sub(
        r"(<span>)(\d+)\s*条消息(</span>)",
        lambda match: f"{match.group(1)}{count} 条消息{match.group(3)}",
        html_text,
        count=1,
    )
    html_text = re.sub(
        r"(id=\"resultCount\">\s*共\s*)(\d+)(\s*条)",
        lambda match: f"{match.group(1)}{count}{match.group(3)}",
        html_text,
        count=1,
    )
    return html_text


def split_messages(messages, chunk_size: int):
    """按固定大小切分消息列表。"""
    for index in range(0, len(messages), chunk_size):
        yield messages[index : index + chunk_size]


def process_html_split(html_path: Path, chunk_size: int):
    """分割 HTML 文件并在同目录输出 part 文件。"""
    html_text = html_path.read_text(encoding="utf-8")
    raw_data, indent, span = extract_data_block(html_text)
    if raw_data is None:
        print(f"❌ 处理失败：在 {html_path.name} 中没有找到聊天数据。")
        return 0, []

    messages = parse_data_list(raw_data)
    if not messages:
        print(f"❌ 处理失败：{html_path.name} 中没有可用消息。")
        return 0, []

    total_parts = (len(messages) + chunk_size - 1) // chunk_size
    base_name = html_path.stem
    split_files = []

    print(f"📋 正在处理 {html_path.name}：共 {len(messages)} 条消息，准备拆分为 {total_parts} 个文件。")

    for index, chunk in enumerate(split_messages(messages, chunk_size), start=1):
        updated_html = update_counts(html_text, len(chunk))
        data_block = build_data_block(chunk, indent)
        updated_html = updated_html[: span[0]] + data_block + updated_html[span[1] :]

        output_name = f"{base_name}_part{index:03d}.html"
        output_path = html_path.parent / output_name
        output_path.write_text(updated_html, encoding="utf-8")
        split_files.append(output_path)
        print(f"  ✅ 已生成：{output_name}（包含 {len(chunk)} 条消息）")

    return total_parts, split_files


def fix_html_print_issue(html_path: Path):
    """修复 HTML 打印问题，生成适合 PDF 打印的副本。"""
    output_html_path = html_path.with_name(f"{html_path.stem}_fixed{html_path.suffix}")
    html_content = html_path.read_text(encoding="utf-8")

    html_content = re.sub(
        r"\.page\s*\{[^}]*height:\s*100vh;[^}]*\}",
        lambda match: match.group(0).replace("height: 100vh;", ""),
        html_content,
    )
    html_content = re.sub(
        r"\.scroll-container\s*\{[^}]*flex:\s*1;[^}]*\}",
        lambda match: match.group(0).replace("flex: 1;", ""),
        html_content,
    )
    html_content = re.sub(
        r"\.scroll-container\s*\{[^}]*min-height:\s*0;[^}]*\}",
        lambda match: match.group(0).replace("min-height: 0;", ""),
        html_content,
    )
    html_content = re.sub(
        r"this\.batchSize\s*=\s*\d+;",
        "this.batchSize = 10000;",
        html_content,
    )
    html_content = re.sub(
        r"\.scroll-container\s*\{[^}]*overflow-y:\s*auto;[^}]*\}",
        lambda match: match.group(0).replace("overflow-y: auto;", "overflow-y: visible;"),
        html_content,
    )

    print_css = """
    <style>
    @media print {
        body { overflow: visible !important; }
        .page { height: auto !important; }
        .scroll-container { overflow: visible !important; height: auto !important; }
        .message-list { height: auto !important; }
        * { overflow: visible !important; height: auto !important; }
    }
    </style>
    """

    html_content = re.sub(r"(</head>)", print_css + r"\1", html_content)
    output_html_path.write_text(html_content, encoding="utf-8")
    return output_html_path


def get_available_browser(playwright):
    """自动检测可用的 Chromium 系浏览器。"""
    browsers_to_try = [
        ("msedge", "Microsoft Edge"),
        ("msedge-beta", "Microsoft Edge Beta"),
        ("msedge-dev", "Microsoft Edge Dev"),
        ("chrome", "Google Chrome"),
        ("chrome-beta", "Google Chrome Beta"),
        ("chromium", "Chromium"),
    ]

    for browser_channel, browser_name in browsers_to_try:
        try:
            print(f"🌐 正在尝试启动浏览器：{browser_name}...")
            browser = playwright.chromium.launch(channel=browser_channel, headless=True)
            print(f"✅ 浏览器可用：{browser_name}")
            return browser, browser_name
        except Exception as exc:
            error_msg = str(exc)
            print(f"❌ {browser_name} 不可用：{error_msg}")
            if "executable" in error_msg.lower() or "not found" in error_msg.lower():
                print(f"   💡 提示：{browser_name} 可能未安装或路径不正确")
            continue

    return None, None


def get_manual_browser_path():
    """获取用户手动输入的浏览器路径。"""
    print("\n" + "=" * 60)
    print("🔍 自动检测浏览器失败，请手动指定浏览器路径")
    print("=" * 60)
    print("💡 常见浏览器路径示例：")
    print("   Microsoft Edge: C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
    print("   Google Chrome: C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
    print("=" * 60)
    
    while True:
        browser_path = input("\n请输入浏览器exe文件的完整路径（或输入 'q' 退出程序）: ").strip().strip('"')
        
        if browser_path.lower() == 'q':
            return None
        
        if not browser_path:
            print("❌ 路径不能为空，请重新输入")
            continue
        
        if not browser_path.endswith('.exe'):
            print("⚠️  警告：路径似乎不是exe文件，请确认")
        
        if not os.path.exists(browser_path):
            print(f"❌ 文件不存在：{browser_path}")
            print("   请检查路径是否正确，或尝试拖拽文件到此处")
            continue
        
        return browser_path


def launch_browser_with_path(playwright, browser_path):
    """使用指定路径启动浏览器。"""
    try:
        print(f"🌐 正在尝试启动浏览器：{browser_path}")
        browser = playwright.chromium.launch(executable_path=browser_path, headless=True)
        print(f"✅ 浏览器启动成功")
        return browser, os.path.basename(browser_path)
    except Exception as exc:
        print(f"❌ 使用指定路径启动浏览器失败：{exc}")
        return None, None


def convert_html_to_pdf(html_files, output_dir: Path):
    """将 HTML 文件转换为 PDF。"""
    print("\n📄 开始把 HTML 转成 PDF，请稍候...")

    browser = None
    context = None
    
    try:
        with sync_playwright() as playwright:
            browser, browser_name = get_available_browser(playwright)
            
            if browser is None:
                print("\n" + "=" * 60)
                print("❌ 自动检测浏览器失败")
                print("=" * 60)
                
                browser_path = get_manual_browser_path()
                
                if browser_path is None:
                    print("\n❌ 用户取消操作，程序退出")
                    return False
                
                browser, browser_name = launch_browser_with_path(playwright, browser_path)
                
                if browser is None:
                    print("\n❌ 无法启动浏览器，转换失败")
                    print("💡 你可以运行以下命令安装浏览器：")
                    print("  - playwright install msedge")
                    print("  - playwright install chrome")
                    print("  - playwright install chromium")
                    return False

            print(f"🚀 本次使用浏览器：{browser_name}")
            context = browser.new_context()

            for html_path in html_files:
                success = convert_single_file_with_timeout(context, html_path, output_dir)
                if not success:
                    print(f"\n⚠️  文件 {html_path.name} 转换失败或被跳过")
                    # 确保资源清理
                    if context:
                        context.close()
                    if browser:
                        browser.close()
                    return False

            print("\n🎉 全部 PDF 转换完成。")
            return True

    except Exception as exc:
        print(f"\n❌ PDF 转换时出现错误：{exc}")
        # 确保异常情况下资源清理
        try:
            if context:
                context.close()
            if browser:
                browser.close()
        except Exception:
            pass
        return False
    finally:
        # 确保资源完全释放
        try:
            if context:
                context.close()
            if browser:
                browser.close()
        except Exception:
            pass


def convert_single_file_with_timeout(context, html_path: Path, output_dir: Path, max_retries: int = 3, timeout_seconds: int = 600):
    """转换单个HTML文件，支持超时和重试机制。"""
    retry_count = 0
    
    while retry_count < max_retries:
        print(f"\n🔄 正在转换：{html_path.name} (尝试 {retry_count + 1}/{max_retries})")
        
        try:
            pdf_path = perform_conversion_with_timeout(context, html_path, output_dir, timeout_seconds)
            print(f"✅ 已完成：{pdf_path.name}")
            return True
        except TimeoutError:
            print(f"⏱️  转换超时（超过 {timeout_seconds // 60} 分钟）")
            retry_count += 1
            
            if retry_count >= max_retries:
                print("\n" + "=" * 60)
                print("❌ 转换时间过长，请减小每个文件包含的消息数量后重试")
                print("=" * 60)
                print("\n请选择接下来怎么操作：")
                print("  1. 重试（重新尝试转换当前文件）")
                print("  2. 跳过（跳过当前文件，继续处理下一个文件）")
                print("  3. 退出（退出程序）")
                
                while True:
                    choice = input("\n请输入选项 (1/2/3): ").strip()
                    if choice == '1':
                        retry_count = 0
                        break
                    elif choice == '2':
                        print(f"⏭️  已跳过：{html_path.name}")
                        return True
                    elif choice == '3':
                        print("👋 程序已退出")
                        return False
                    else:
                        print("❌ 无效选项，请重新输入")
            else:
                print(f"🔄 正在进行第 {retry_count + 1} 次重试...")
        except Exception as exc:
            print(f"❌ 转换失败：{exc}")
            retry_count += 1
            if retry_count >= max_retries:
                print(f"❌ 文件 {html_path.name} 转换失败，已达到最大重试次数")
                return False
            print(f"🔄 正在进行第 {retry_count + 1} 次重试...")
    
    return False


def perform_conversion_with_timeout(context, html_path: Path, output_dir: Path, timeout_seconds: int):
    """执行实际的PDF转换操作，使用Playwright内置超时。"""
    page = None
    try:
        page = context.new_page()
        
        page.goto(html_path.resolve().as_uri(), timeout=timeout_seconds * 1000)
        
        pdf_bytes = page.pdf(
            format="A4",
            print_background=True,
            prefer_css_page_size=True
        )

        pdf_path = output_dir / f"{html_path.stem}.pdf"
        pdf_path.write_bytes(pdf_bytes)
        
        return pdf_path
    finally:
        if page:
            try:
                page.close()
            except Exception:
                pass


def wait_for_key():
    """等待用户按任意键继续。"""
    print("\n👋 操作结束，按任意键退出程序...")
    try:
        if os.name == "nt":
            import msvcrt

            msvcrt.getch()
        else:
            import termios
            import tty

            fd = sys.stdin.fileno()
            old_settings = termios.tcgetattr(fd)
            try:
                tty.setraw(fd)
                sys.stdin.read(1)
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
    except Exception:
        input("👋 按回车键退出程序...")


def process_directory(directory_path: Path, fixed_chunk_size: int = None):
    """处理单个目录中的HTML文件。
    
    Args:
        directory_path: 要处理的目录路径
        fixed_chunk_size: 固定的消息数量，如果为None则询问用户
    """
    print(f"\n📂 正在处理目录：{directory_path}")
    
    # 切换到目标目录
    original_cwd = Path.cwd()
    os.chdir(directory_path)
    
    html_files = sorted([path for path in directory_path.glob("*.html") if path.is_file()])
    if not html_files:
        print("❌ 该目录下没有html文件，跳过处理")
        os.chdir(original_cwd)
        return False

    print(f"📁 已找到 {len(html_files)} 个 HTML 文件：")
    for html_file in html_files:
        print(f"  📄 {html_file.name}")

    # 确定消息数量
    config_path = directory_path / "config.ini"
    if config_path.exists():
        chunk_size = load_cut_size(config_path)
        print(f"\n📋 已从配置文件读取拆分大小：每个文件 {chunk_size} 条消息")
    elif fixed_chunk_size is not None:
        # 使用固定的消息数量（批量处理模式）
        chunk_size = fixed_chunk_size
        print(f"\n📋 批量处理模式：每个文件固定包含 {chunk_size} 条消息")
    else:
        # 询问用户输入（单目录处理模式）
        try:
            chunk_size = int(input("\n请输入每个文件包含多少条消息（默认4000）：") or "4000")
            chunk_size = max(1, chunk_size)
        except ValueError:
            chunk_size = 4000
            print(f"⚠️  输入无效，已使用默认值：{chunk_size}")

    pdf_dir = directory_path / "PDF输出"
    temp_files = []

    print("\n" + "=" * 40)
    print("🔪 第1步：拆分 HTML 文件")
    print("=" * 40)

    total_parts = 0
    all_split_files = []
    for html_file in html_files:
        parts, split_files = process_html_split(html_file, chunk_size)
        total_parts += parts
        all_split_files.extend(split_files)

    if total_parts == 0:
        print("❌ 没有成功拆分任何文件，跳过该目录。")
        os.chdir(original_cwd)
        return False

    print("\n" + "=" * 40)
    print("🔧 第2步：修复 HTML 打印格式")
    print("=" * 40)

    fixed_files = []
    for split_file in all_split_files:
        print(f"🔧 正在修复：{split_file.name}")
        fixed_file = fix_html_print_issue(split_file)
        fixed_files.append(fixed_file)
        print(f"  ✅ 修复完成：{fixed_file.name}")

    print("\n" + "=" * 40)
    print("📄 第3步：转换为 PDF(此过程可能较慢，请耐心等待...)")
    print("=" * 40)

    pdf_dir.mkdir(exist_ok=True)
    success = convert_html_to_pdf(fixed_files, pdf_dir)

    print("\n" + "=" * 40)
    print("🧹 第4步：清理临时文件")
    print("=" * 40)

    temp_files.extend(all_split_files)
    temp_files.extend(fixed_files)

    deleted_count = 0
    for temp_file in temp_files:
        try:
            if temp_file.exists():
                temp_file.unlink()
                deleted_count += 1
                print(f"  🗑️  已删除：{temp_file.name}")
        except Exception as exc:
            print(f"  ❌ 删除失败 {temp_file.name}：{exc}")

    print(f"🧹 已清理 {deleted_count} 个临时文件。")

    print("\n" + "=" * 60)
    print("📊 处理完成总结")
    print("=" * 60)
    print(f"📄 原始 HTML 文件：{len(html_files)} 个")
    print(f"🔪 拆分后文件：{total_parts} 个")
    print(f"🔧 修复后文件：{len(fixed_files)} 个")
    print(f"📄 PDF 输出文件：{len(list(pdf_dir.glob('*.pdf')))} 个")
    print(f"🧹 清理临时文件：{deleted_count} 个")

    if success:
        print("\n🎉 该目录处理完成。")
        print(f"📄 最终保留：{len(html_files)} 个原始 HTML 文件")
        print(f"📄 最终保留：{len(list(pdf_dir.glob('*.pdf')))} 个 PDF 文件")
        print("🧹 所有临时文件已自动清理。")
    else:
        print("\n❌ PDF 转换可能存在问题，请检查浏览器是否安装。")

    # 恢复原始工作目录
    os.chdir(original_cwd)
    
    # 强制垃圾回收，确保浏览器资源完全释放
    import gc
    gc.collect()
    
    # 添加短暂延迟，确保系统资源完全释放
    time.sleep(1)
    
    return success


def main():
    """主函数。"""
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent

    print("🚀 聊天记录html转PDF插件启动")
    print("📞 技术支持联系方式：dawn10370108")
    print("=" * 60)
    print(f"📁 当前工作目录：{base_dir}")

    # 选择处理模式
    print("\n请选择处理模式：")
    print("  1. 处理当前目录（直接识别当前目录下的HTML文件）")
    print("  2. 批量处理多个目录（手动填写总目录路径，自动识别子目录）")
    
    while True:
        choice = input("\n请输入选项 (1/2): ").strip()
        if choice == '1':
            # 处理当前目录
            success = process_directory(base_dir)
            break
        elif choice == '2':
            # 批量处理多个目录
            print("\n" + "=" * 60)
            print("📂 批量处理模式：请输入总目录路径")
            print("=" * 60)
            print("💡 程序会自动识别该目录下的所有子文件夹，并在每个子文件夹中执行转换")
            print("   批量处理模式下，所有子文件夹将使用相同的消息数量设置")
            print("   例如：输入 C:\\聊天记录，程序会处理 C:\\聊天记录\\文件夹1、C:\\聊天记录\\文件夹2 等")
            
            while True:
                root_dir_path = input("\n请输入总目录路径（或输入 'q' 返回主菜单）: ").strip().strip('"')
                
                if root_dir_path.lower() == 'q':
                    print("👋 已返回主菜单")
                    wait_for_key()
                    return
                
                if not root_dir_path:
                    print("❌ 路径不能为空，请重新输入")
                    continue
                
                root_dir = Path(root_dir_path)
                if not root_dir.exists():
                    print(f"❌ 目录不存在：{root_dir}")
                    print("   请检查路径是否正确，或尝试拖拽文件夹到此处")
                    continue
                
                if not root_dir.is_dir():
                    print(f"❌ 路径不是目录：{root_dir}")
                    continue
                
                # 查找所有子目录
                sub_dirs = [d for d in root_dir.iterdir() if d.is_dir()]
                if not sub_dirs:
                    print(f"❌ 该目录下没有子文件夹：{root_dir}")
                    continue
                
                print(f"\n📁 找到 {len(sub_dirs)} 个子文件夹：")
                for sub_dir in sub_dirs:
                    print(f"  📂 {sub_dir.name}")
                
                # 批量处理模式下统一询问消息数量
                print("\n" + "=" * 40)
                print("📋 批量处理设置：请输入每个文件包含多少条消息")
                print("=" * 40)
                
                try:
                    batch_chunk_size = int(input("\n请输入每个文件包含多少条消息（默认4000）：") or "4000")
                    batch_chunk_size = max(1, batch_chunk_size)
                    print(f"✅ 已设置：每个文件包含 {batch_chunk_size} 条消息")
                except ValueError:
                    batch_chunk_size = 4000
                    print(f"⚠️  输入无效，已使用默认值：{batch_chunk_size}")
                
                confirm = input("\n确认开始批量处理这些文件夹吗？(y/n): ").strip().lower()
                if confirm == 'y' or confirm == 'yes' or confirm == '是':
                    total_success = 0
                    for i, sub_dir in enumerate(sub_dirs, 1):
                        print(f"\n{'='*60}")
                        print(f"📂 正在处理第 {i}/{len(sub_dirs)} 个文件夹：{sub_dir.name}")
                        print('='*60)
                        
                        if process_directory(sub_dir, batch_chunk_size):
                            total_success += 1
                        
                        if i < len(sub_dirs):
                            print("\n⏭️  继续处理下一个文件夹...")
                    
                    print(f"\n🎉 批量处理完成！成功处理 {total_success}/{len(sub_dirs)} 个文件夹")
                    break
                else:
                    print("👋 已取消批量处理")
                    continue
            
            break
        else:
            print("❌ 无效选项，请重新输入")

    wait_for_key()

    print(f"\n📁 已找到 {len(html_files)} 个 HTML 文件：")
    for html_file in html_files:
        print(f"  📄 {html_file.name}")

    config_path = base_dir / "config.ini"
    if config_path.exists():
        chunk_size = load_cut_size(config_path)
        print(f"\n📋 已从配置文件读取拆分大小：每个文件 {chunk_size} 条消息")
    else:
        try:
            chunk_size = int(input("\n请输入每个文件包含多少条消息（默认4000）：") or "4000")
            chunk_size = max(1, chunk_size)
        except ValueError:
            chunk_size = 4000
            print(f"⚠️  输入无效，已使用默认值：{chunk_size}")

    pdf_dir = base_dir / "PDF输出"
    temp_files = []

    print("\n" + "=" * 40)
    print("🔪 第1步：拆分 HTML 文件")
    print("=" * 40)

    total_parts = 0
    all_split_files = []
    for html_file in html_files:
        parts, split_files = process_html_split(html_file, chunk_size)
        total_parts += parts
        all_split_files.extend(split_files)

    if total_parts == 0:
        print("❌ 没有成功拆分任何文件，程序将退出。")
        wait_for_key()
        return

    print("\n" + "=" * 40)
    print("🔧 第2步：修复 HTML 打印格式")
    print("=" * 40)

    fixed_files = []
    for split_file in all_split_files:
        print(f"🔧 正在修复：{split_file.name}")
        fixed_file = fix_html_print_issue(split_file)
        fixed_files.append(fixed_file)
        print(f"  ✅ 修复完成：{fixed_file.name}")

    print("\n" + "=" * 40)
    print("📄 第3步：转换为 PDF(此过程可能较慢，请耐心等待...)")
    print("=" * 40)

    pdf_dir.mkdir(exist_ok=True)
    success = convert_html_to_pdf(fixed_files, pdf_dir)

    print("\n" + "=" * 40)
    print("🧹 第4步：清理临时文件")
    print("=" * 40)

    temp_files.extend(all_split_files)
    temp_files.extend(fixed_files)

    deleted_count = 0
    for temp_file in temp_files:
        try:
            if temp_file.exists():
                temp_file.unlink()
                deleted_count += 1
                print(f"  🗑️  已删除：{temp_file.name}")
        except Exception as exc:
            print(f"  ❌ 删除失败 {temp_file.name}：{exc}")

    print(f"🧹 已清理 {deleted_count} 个临时文件。")

    print("\n" + "=" * 60)
    print("📊 处理完成总结")
    print("=" * 60)
    print(f"📄 原始 HTML 文件：{len(html_files)} 个")
    print(f"🔪 拆分后文件：{total_parts} 个")
    print(f"🔧 修复后文件：{len(fixed_files)} 个")
    print(f"📄 PDF 输出文件：{len(list(pdf_dir.glob('*.pdf')))} 个")
    print(f"🧹 清理临时文件：{deleted_count} 个")

    if success:
        print("\n🎉 全部步骤已完成。")
        print(f"📄 最终保留：{len(html_files)} 个原始 HTML 文件")
        print(f"📄 最终保留：{len(list(pdf_dir.glob('*.pdf')))} 个 PDF 文件")
        print("🧹 所有临时文件已自动清理。")
    else:
        print("\n❌ PDF 转换可能存在问题，请检查浏览器是否安装。")

    wait_for_key()


if __name__ == "__main__":
    main()