import os
import datetime
from docx import Document  # Word 문서 저장용 모듈

# --- 파일을 읽어서 문자열로 반환하는 함수 ---
def load_file(filepath):
    """
    주어진 파일 경로를 받아 파일을 열고, 내용을 UTF-8로 읽어서 반환한다.
    파일이 없으면 FileNotFoundError가 발생하므로 이를 처리해 사용자에게 알리고 None을 반환한다.
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(f"[오류] 파일을 찾을 수 없습니다: {filepath}")
        return None

# --- 폴더 내 파일 목록을 출력하고 리스트로 반환하는 함수 ---
def list_files(folder):
    """
    지정한 폴더에서 파일 목록을 읽어와 사용자에게 번호와 함께 출력한다.
    폴더가 없거나 파일이 하나도 없으면 적절한 안내 메시지를 출력하고 빈 리스트를 반환한다.
    """
    try:
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        if not files:
            print(f"[안내] '{folder}' 폴더에 파일이 없습니다.")
            return []
        print("\n선택 가능한 파일 목록:")
        for i, f in enumerate(files, 1):
            print(f"{i}. {f}")
        return files
    except FileNotFoundError:
        print(f"[오류] 폴더 '{folder}'를 찾을 수 없습니다.")
        return []

# --- 텍스트를 문장 단위로 분리하고 각 문장 끝에 '(O/X)' 붙이는 함수 ---
def generate_ox_quiz(text):
    """
    입력된 텍스트를 온점(.)을 기준으로 문장으로 분리한다.
    각 문장 끝에 ' (O/X)'를 붙여서 리스트로 반환한다.
    빈 문장이나 공백 문장은 제외한다.
    """
    sentences = [s.strip() for s in text.split('.') if s.strip()]
    return [s + '. (O/X)' for s in sentences]

# --- 퀴즈 목록을 Word(.docx) 파일로 저장하는 함수 ---
def save_to_docx(quiz_items, filename="quiz_output.docx"):
    """
    퀴즈 문장 리스트를 받아 새 Word 문서로 저장한다.
    문서 제목을 추가하고 각 퀴즈 문장을 번호와 함께 문단으로 삽입한다.
    저장이 완료되면 사용자에게 저장된 파일명을 출력한다.
    """
    doc = Document()
    doc.add_heading('O/X 퀴즈 목록', level=1)
    for i, item in enumerate(quiz_items, 1):
        doc.add_paragraph(f"{i}. {item}")
    doc.save(filename)
    print(f"퀴즈가 Word 파일로 저장되었습니다: {filename}")

# --- Y/N 입력을 받을 때 유효성 검사하며 묻는 함수 ---
def ask_yes_no(prompt):
    """
    사용자에게 Y 또는 N만 입력 받도록 반복해서 묻는 함수.
    Y이면 True, N이면 False 반환.
    잘못된 입력 시 안내 메시지 출력 후 다시 질문.
    """
    while True:
        ans = input(prompt).strip().lower()
        if ans == 'y':
            return True
        elif ans == 'n':
            return False
        else:
            print("[경고] Y 또는 N만 입력해주세요.")

# --- 메인 프로그램 흐름을 담당하는 함수 ---
def main():
    folder = "study_files"

    while True:
        files = list_files(folder)
        if not files:
            print("프로그램을 종료합니다.")
            break

        print("\n0. 종료하기")  # 0 입력 시 종료 안내 추가

        try:
            choice = int(input("불러올 파일 번호를 입력하세요 (0 입력 시 종료): "))
            if choice == 0:
                print("프로그램을 종료합니다.")
                break
            choice -= 1  # 리스트 인덱스 맞춤
            if choice < 0 or choice >= len(files):
                print("[경고] 잘못된 번호입니다.\n")
                continue
        except ValueError:
            print("[경고] 숫자를 입력해야 합니다.\n")
            continue

        filename = files[choice]
        filepath = os.path.join(folder, filename)

        if not filename.endswith(('.txt', '.md')):
            print("[경고] 지원하지 않는 파일 형식입니다.\n")
            continue

        text = load_file(filepath)
        if text:
            quiz_items = generate_ox_quiz(text)
            print("\nO/X 퀴즈 목록:\n")
            for i, q in enumerate(quiz_items, 1):
                print(f"{i}. {q}")

            if ask_yes_no("\n현재 퀴즈를 Word(.docx) 파일로 저장하시겠습니까? (Y/N): "):
                docx_name = input("저장할 파일 이름을 입력하세요 (확장자 제외, 엔터 누르면 자동 생성): ").strip()
                if not docx_name:
                    now = datetime.datetime.now()
                    docx_name = now.strftime("quiz_%Y%m%d_%H%M%S")
                try:
                    save_to_docx(quiz_items, f"{docx_name}.docx")
                except Exception as e:
                    print(f"[오류] Word 저장에 실패했습니다: {e}")
        else:
            print("[오류] 파일을 불러오지 못했습니다.\n")
            continue

        if not ask_yes_no("\n프로그램을 다시 실행하시겠습니까? (Y/N): "):
            print("프로그램을 종료합니다. 감사합니다.")
            break

# --- 프로그램 실행 진입점 ---
if __name__ == "__main__":
    main()
