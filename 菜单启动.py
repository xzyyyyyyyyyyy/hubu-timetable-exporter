import sys


def pause():
    input("\n按回车返回菜单...")


def run_task(task_func, name):
    try:
        task_func()
    except SystemExit:
        raise
    except Exception as exc:
        print(f"[{name}] 运行出错：{exc}")
    finally:
        pause()


def main():
    while True:
        print("\n=== 湖北大学课表工具 ===")
        print("常用功能：")
        print("3. 导出课表（单个班级）")
        print("4. 导出课表（批量）")
        print("高级功能：")
        print("1. 采集课表参数（仅维护/调试）")
        print("2. 筛选课表参数（仅维护/调试）")
        print("0. 退出")

        choice = input("请输入序号：").strip()
        if choice == "1":
            import 采集参数
            run_task(采集参数.main, "采集参数")
        elif choice == "2":
            import 筛选课表参数
            run_task(筛选课表参数.main, "筛选课表参数")
        elif choice == "3":
            import 选择采集课表
            run_task(选择采集课表.main, "选择采集课表")
        elif choice == "4":
            import 批量采集课表
            run_task(批量采集课表.main, "批量采集课表")
        elif choice == "0":
            print("已退出。")
            sys.exit(0)
        else:
            print("输入有误，请重新输入。")


if __name__ == "__main__":
    main()
