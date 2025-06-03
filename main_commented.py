import os
import datetime
from docx import Document  # Word 문서 저장용 모듈

# --- 파일을 읽어서 문자열로 반환하는 함수 ---
def load_file(filepath):
    """
    주어진 파일 경로를 받아 파일을 열고, 내용을 UTF-8로 읽어서 반환한다.
    UTF-8은 한글 등 다양한 문자를 지원하는 인코딩 방식이다.
    파일이 없으면 FileNotFoundError 예외가 발생하며, 이를 처리해 사용자에게 알리고 None을 반환한다.
    None 반환은 호출하는 함수가 파일 읽기 실패를 인지할 수 있게 한다.
    """
    try:
        # 'r' 모드로 파일을 열되, UTF-8 인코딩으로 읽음
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()  # 파일 내용 전체를 문자열로 읽어 반환
    except FileNotFoundError:
        # 파일이 존재하지 않을 경우 경고 메시지 출력
        print(f"[오류] 파일을 찾을 수 없습니다: {filepath}")
        return None

# --- 폴더 내 파일 목록을 출력하고 리스트로 반환하는 함수 ---
def list_files(folder):
    """
    지정한 폴더에서 모든 파일을 찾아 리스트로 반환하고, 사용자에게 번호와 함께 출력한다.
    폴더가 없거나 파일이 하나도 없으면 알림 메시지를 출력하고 빈 리스트를 반환한다.
    """
    try:
        # os.listdir은 폴더 내 모든 항목을 반환하므로,
        # os.path.isfile으로 실제 파일만 필터링한다.
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        if not files:
            # 파일이 없으면 사용자에게 알려줌
            print(f"[안내] '{folder}' 폴더에 파일이 없습니다.")
            return []
        print("\n선택 가능한 파일 목록:")
        # 번호 붙여서 파일명 출력 (1부터 시작)
        for i, f in enumerate(files, 1):
            print(f"{i}. {f}")
        return files  # 파일 리스트 반환
    except FileNotFoundError:
        # 폴더가 존재하지 않으면 오류 메시지 출력
        print(f"[오류] 폴더 '{folder}'를 찾을 수 없습니다.")
        return []

# --- 텍스트를 문장 단위로 분리하고 각 문장 끝에 '(O/X)' 붙이는 함수 ---
def generate_ox_quiz(text):
    """
    입력된 긴 텍스트를 온점(.)을 기준으로 문장별로 나눈다.
    빈 문장이나 공백 문장은 제외한다.
    각 문장 끝에 '. (O/X)'를 붙여 학습용 O/X 퀴즈 문장 리스트를 반환한다.
    """
    # 문장 분리 후 각 문장 양끝 공백 제거
    sentences = [s.strip() for s in text.split('.') if s.strip()]
    # 각 문장 끝에 O/X 표시를 붙여서 리스트 생성
    return [s + '. (O/X)' for s in sentences]

# --- 퀴즈 목록을 Word(.docx) 파일로 저장하는 함수 ---
def save_to_docx(quiz_items, filename="quiz_output.docx"):
    """
    퀴즈 문장 리스트를 받아 새 Word 문서를 생성한다.
    문서 상단에 제목 'O/X 퀴즈 목록'을 추가하고,
    각 문장을 번호와 함께 단락으로 추가한다.
    지정한 이름으로 파일을 저장한 후 완료 메시지를 출력한다.
    """
    doc = Document()  # 새 문서 객체 생성
    doc.add_heading('O/X 퀴즈 목록', level=1)  # 1단계 제목 추가
    for i, item in enumerate(quiz_items, 1):
        # 각 퀴즈 문장을 번호 붙여 문단으로 추가
        doc.add_paragraph(f"{i}. {item}")
    doc.save(filename)  # 파일 저장
    print(f"퀴즈가 Word 파일로 저장되었습니다: {filename}")

# --- Y/N 입력을 받는 함수 ---
def ask_yes_no(prompt):
    """
    사용자에게 Y 또는 N만 입력받도록 반복해서 묻는 함수다.
    올바른 입력이 들어올 때까지 계속 질문한다.
    Y는 True, N은 False를 반환한다.
    다른 입력이 들어오면 경고 메시지를 출력한다.
    """
    while True:
        ans = input(prompt).strip().lower()  # 입력 받은 문자열을 소문자로 변환 후 양끝 공백 제거
        if ans == 'y':
            return True  # Y 입력 시 True 반환
        elif ans == 'n':
            return False  # N 입력 시 False 반환
        else:
            # Y 또는 N이 아닌 다른 입력 시 안내 메시지 출력
            print("[경고] Y 또는 N만 입력해주세요.")

# --- 메인 프로그램 흐름 ---
def main():
    folder = "study_files"  # 퀴즈 원본 텍스트 파일이 저장된 폴더 이름

    while True:  # 사용자 종료 전까지 반복
        files = list_files(folder)  # 폴더 내 파일 목록 가져오기
        if not files:
            print("프로그램을 종료합니다.")  # 파일이 없으면 프로그램 종료
            break

        print("\n0. 종료하기")  # 사용자에게 종료 옵션 안내

        try:
            choice = int(input("불러올 파일 번호를 입력하세요 (0 입력 시 종료): "))
            if choice == 0:
                print("프로그램을 종료합니다.")  # 0 입력 시 종료
                break
            choice -= 1  # 0부터 시작하는 인덱스로 변환
            if choice < 0 or choice >= len(files):
                print("[경고] 잘못된 번호입니다.\n")  # 범위 벗어난 번호 입력 시 경고
                continue  # 다시 입력 받음
        except ValueError:
            print("[경고] 숫자를 입력해야 합니다.\n")  # 숫자가 아닌 값 입력 시 경고
            continue  # 다시 입력 받음

        filename = files[choice]  # 선택한 파일명
        filepath = os.path.join(folder, filename)  # 파일 경로 생성

        # 지원하지 않는 파일 형식 걸러내기 (txt, md만 허용)
        if not filename.endswith(('.txt', '.md')):
            print("[경고] 지원하지 않는 파일 형식입니다.\n")
            continue  # 다시 입력 받음

        text = load_file(filepath)  # 파일 내용 읽기
        if text:
            quiz_items = generate_ox_quiz(text)  # 문장 단위로 나누고 O/X 붙이기
            print("\nO/X 퀴즈 목록:\n")
            for i, q in enumerate(quiz_items, 1):
                print(f"{i}. {q}")  # 콘솔에 퀴즈 출력

            # Word 저장 여부 사용자에게 물어보기
            if ask_yes_no("\n현재 퀴즈를 Word(.docx) 파일로 저장하시겠습니까? (Y/N): "):
                docx_name = input("저장할 파일 이름을 입력하세요 (확장자 제외, 엔터 누르면 자동 생성): ").strip()
                if not docx_name:
                    # 입력 없으면 현재 날짜시간 기준 자동 파일명 생성
                    now = datetime.datetime.now()
                    docx_name = now.strftime("quiz_%Y%m%d_%H%M%S")
                try:
                    save_to_docx(quiz_items, f"{docx_name}.docx")  # 저장 시도
                except Exception as e:
                    print(f"[오류] Word 저장에 실패했습니다: {e}")
        else:
            print("[오류] 파일을 불러오지 못했습니다.\n")
            continue  # 다시 입력 받음

        # 프로그램 재실행 여부 묻기
        if not ask_yes_no("\n프로그램을 다시 실행하시겠습니까? (Y/N): "):
            print("프로그램을 종료합니다. 감사합니다.")  # 종료 인사
            break

# --- 프로그램 시작 ---
if __name__ == "__main__":
    main()
