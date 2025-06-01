import os
from docx import Document  # Word 파일 저장용 모듈 사용

# 파일을 읽어서 문자열로 반환하는 함수
def load_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(f"[오류] 파일을 찾을 수 없습니다: {filepath}")
        return None

# 폴더 내 파일 목록을 출력하고 리스트로 반환하는 함수
def list_files(folder):
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

# 텍스트에서 온점(.) 기준으로 문장을 나눈 후 (O/X) 를 붙이는 함수
def generate_ox_quiz(text):
    sentences = [s.strip() for s in text.split('.') if s.strip()]  # 공백 제거 + 빈 문장 제외
    return [s + '. (O/X)' for s in sentences]  # 각 문장 끝에 (O/X) 붙이기

# 퀴즈 목록을 Word(.docx) 파일로 저장하는 함수
def save_to_docx(quiz_items, filename="quiz_output.docx"):
    doc = Document()
    doc.add_heading('O/X 퀴즈 목록', level=1)  # 문서 제목 설정
    for i, item in enumerate(quiz_items, 1):
        doc.add_paragraph(f"{i}. {item}")  # 번호 + 퀴즈 문장 추가
    doc.save(filename)
    print(f"퀴즈가 Word 파일로 저장되었습니다: {filename}")

# 메인 프로그램 흐름을 담당하는 함수
def main():
    folder = "study_files"

    while True:
        files = list_files(folder)  # 파일 목록 불러오기
        if not files:
            print("프로그램을 종료합니다.")
            break

        try:
            choice = int(input("\n불러올 파일 번호를 입력하세요: ")) - 1
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

            # 사용자에게 Word 저장 여부 묻기
            save_docx = input("\n현재 퀴즈를 Word(.docx) 파일로 저장하시겠습니까? (Y/N): ").strip().lower()
            if save_docx == 'y':
                docx_name = input("저장할 파일 이름을 입력하세요 (확장자 제외): ").strip()
                if not docx_name:
                    docx_name = "quiz_output"
                try:
                    save_to_docx(quiz_items, f"{docx_name}.docx")
                except Exception as e:
                    print(f"[오류] Word 저장에 실패했습니다: {e}")
        else:
            print("[오류] 파일을 불러오지 못했습니다.\n")
            continue

        # 사용자에게 프로그램 반복 실행 여부 묻기
        again = input("\n프로그램을 다시 실행하시겠습니까? (Y/N): ").strip().lower()
        if again != 'y':
            print("프로그램을 종료합니다. 감사합니다.")
            break

# 프로그램 실행 진입점
if __name__ == "__main__":
    main()
