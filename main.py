from __future__ import annotations

import json
import sys
from pathlib import Path

# ANSI 颜色码
class Colors:
    USER = "\033[96m"      # 青色 - 用户输入
    ASSISTANT = "\033[92m" # 绿色 - AI 输出
    RESET = "\033[0m"       # 重置
    BOLD = "\033[1m"
    DIM = "\033[2m"
    MAGENTA = "\033[95m"   # 紫红色 - 装饰

if __package__:
    from .graph import OfficeAgent
    from .state import create_initial_state
else:
    CURRENT_DIR = Path(__file__).resolve().parent
    PARENT_DIR = CURRENT_DIR.parent
    if str(PARENT_DIR) not in sys.path:
        sys.path.insert(0, str(PARENT_DIR))

    from office_agent.graph import OfficeAgent
    from office_agent.state import create_initial_state


def parse_response(response: str) -> tuple[str, str]:
    """解析 AI 返回的结果，分离内容和思考过程"""
    try:
        data = json.loads(response)
        return data.get("content", ""), data.get("reasoning", "")
    except json.JSONDecodeError:
        return response, ""


def main() -> None:
    agent = OfficeAgent()

    # 欢迎信息 - 简洁版
    print(f"{Colors.BOLD}{Colors.MAGENTA}")
    print("  ╭─────────────────────────────────────────╮")
    print("  │      Office Agent  ·  自动化办公助手     │")
    print("  ╰─────────────────────────────────────────╯")
    print(f"{Colors.DIM}  输入 quit 退出{Colors.RESET}")

    while True:
        try:
            user_input = input(f"\n{Colors.USER}▶{Colors.RESET} ").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{Colors.DIM}已退出。{Colors.RESET}")
            break
        if not user_input:
            continue
        if user_input.lower() in {"quit", "exit"}:
            print(f"{Colors.DIM}再见！{Colors.RESET}")
            break

        result = agent.invoke(create_initial_state(user_input))
        response_text = result.get("response", "")

        # 解析思考过程（不显示）
        content, _ = parse_response(response_text)

        # 显示 AI 回复
        print(f"\n{Colors.ASSISTANT}◇ {content}{Colors.RESET}")

        if result.get("error"):
            print(f"{Colors.DIM}⚠ {result['error']}{Colors.RESET}")


if __name__ == "__main__":
    main()
