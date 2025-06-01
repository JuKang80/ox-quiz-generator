import os

# 파일을 읽어 텍스트 반환
def load_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(f"파일을 찾을 수 없습니다: {filepath}")
        return None
    except Exception as e:
        print(f"파일을 여는 중 오류 발생: {e}")
        return None

# 폴더 내 파일 목록 출력
def list_files(folder):
    try:
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        if not files:
            print(f"'{folder}' 폴더에 파일이 없습니다.")
            return []
        print("파일 목록:")
        for i, f in enumerate(files, 1):
            print(f"{i}. {f}")
        return files
    except FileNotFoundError:
        print(f"폴더 '{folder}'를 찾을 수 없습니다.")
        return []

# 온점 기준으로 문장 분리하고 (O/X) 붙이기
def generate_ox_quiz(text):
    sentences = [s.strip() for s in text.split('.') if s.strip()]
    return [s + '. (O/X)' for s in sentences]

# 실행 흐름
def main():
    folder = "study_files"
    files = list_files(folder)
    if not files:
        return

    try:
        choice = int(input("불러올 파일 번호를 입력하세요: ")) - 1
        if choice < 0 or choice >= len(files):
            print("잘못된 번호입니다.")
            return
    except ValueError:
        print("숫자를 입력해야 합니다.")
        return

    filename = files[choice]
    filepath = os.path.join(folder, filename)

    if not filename.endswith(('.txt', '.md')):
        print("사용할 수 없는 파일입니다. 다시 시도해주세요.")
        return

    text = load_file(filepath)
    if text:
        quiz_items = generate_ox_quiz(text)
        print("\nO/X 퀴즈 목록:\n")
        for i, q in enumerate(quiz_items, 1):
            print(f"{i}. {q}")
    else:
        print("파일을 불러오지 못했습니다.")

# 실행 진입점
if __name__ == "__main__":
    main()
