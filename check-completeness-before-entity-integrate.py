import os
import string

def ask_int(prompt, min_value=None):
    while True:
        try:
            v = int(input(prompt))
            if min_value is not None and v < min_value:
                raise ValueError
            return v
        except ValueError:
            print("❌ 請輸入有效的整數")

def main():
    print("🔍 檔案完整性檢查工具\n")

    base_path = input("📁 請輸入欲檢驗的資料夾路徑（例如 ./Germany）: ").strip()
    if not os.path.isdir(base_path):
        print("❌ 路徑不存在")
        return

    country = input("🏳️ 國家（例如 Germany）: ").strip()

    entity_count = ask_int("👥 entity 數量: ", min_value=1)
    start_year = ask_int("📆 開始年: ")
    end_year = ask_int("📆 結束年: ")

    if start_year > end_year:
        print("❌ 開始年不可大於結束年")
        return

    group_count = ask_int("🧩 變數分組數: ", min_value=1)

    ext_input = input("📄 副檔名（預設 xlsx,xlsm，直接 Enter 使用預設）: ").strip()
    if ext_input:
        extensions = tuple(e.strip().lstrip(".") for e in ext_input.split(","))
    else:
        extensions = ("xlsx", "xlsm")

    # ===== 開始檢查 =====
    groups = list(string.ascii_uppercase[:group_count])

    existing_files = set()
    for fname in os.listdir(base_path):
        name, ext = os.path.splitext(fname)
        if ext.lstrip(".") in extensions:
            existing_files.add(name)

    missing = []

    for entity in range(1, entity_count + 1):
        for year in range(start_year, end_year + 1):
            for g in groups:
                fname = f"{country}{entity}-{year}{g}"
                if fname not in existing_files:
                    missing.append(fname)

    # ===== 輸出結果 =====
    total_expected = entity_count * (end_year - start_year + 1) * group_count

    print("\n" + "=" * 60)
    print("📊 檢查結果")
    print("=" * 60)
    print(f"✅ 應有檔案數: {total_expected}")
    print(f"📦 實際檔案數: {len(existing_files)}")
    print(f"❌ 缺失檔案數: {len(missing)}")

    if missing:
        print("\n🚨 缺失檔案列表:")
        for f in missing:
            print("  -", f)
    else:
        print("\n🎉 沒有缺檔，資料完整！")

if __name__ == "__main__":
    main()
