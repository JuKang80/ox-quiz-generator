import os
from docx import Document  # Word 문서 생성 및 저장에 필요한 외부 모듈

# --- 파일을 열어 텍스트를 읽고 문자열로 반환하는 함수 ---
def load_file(filepath):
    """
    주어진 파일 경로를 받아 파일을 열고, 내용을 UTF-8로 읽어서 반환한다.
    파일이 없으면 FileNotFoundError가 발생하므로 이를 처리해 사용자에게 알리고 None을 반환한다.
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:  # UTF-8 인코딩으로 파일 열기
            return f.read()  # 파일 전체 내용을 문자열로 반환
    except FileNotFoundError:
        print(f"[오류] 파일을 찾을 수 없습니다: {filepath}")  # 오류 메시지 출력
        return None  # 호출자가 None인지 체크할 수 있도록 반환

# --- 폴더 내 파일 목록을 출력하고 리스트로 반환하는 함수 ---
def list_files(folder):
    """
    지정된 폴더에서 파일 목록을 읽어와 사용자에게 번호와 함께 출력한다.
    폴더가 없거나 파일이 없으면 적절한 안내 메시지를 출력하고 빈 리스트를 반환한다.
    """
    try:
        # os.listdir은 폴더 내 모든 항목 반환, os.path.isfile으로 파일만 필터링
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        if not files:
            print(f"[안내] '{folder}' 폴더에 파일이 없습니다.")  # 파일이 없으면 안내
            return []  # 빈 리스트 반환으로 호출자가 알 수 있도록 함
        print("\n선택 가능한 파일 목록:")  # 사용자 안내 메시지
        for i, f in enumerate(files, 1):  # 번호 붙여서 출력
            print(f"{i}. {f}")
        return files  # 파일 리스트 반환
    except FileNotFoundError:
        print(f"[오류] 폴더 '{folder}'를 찾을 수 없습니다.")  # 폴더가 없으면 에러 안내
        return []  # 빈 리스트 반환

# --- 텍스트를 문장 단위로 분리하고 각 문장 끝에 '(O/X)' 붙이는 함수 ---
def generate_ox_quiz(text):
    """
    입력된 텍스트를 온점(.)을 기준으로 문장으로 분리한다.
    각 문장 끝에 ' (O/X)'를 붙여서 리스트로 반환한다.
    빈 문장이나 공백 문장은 제외한다.
    """
    sentences = [s.strip() for s in text.split('.') if s.strip()]  # 공백 제거 및 빈 문장 제외
    return [s + '. (O/X)' for s in sentences]  # 각 문장에 O/X 표시 추가

# --- 퀴즈 목록을 Word 문서(.docx)로 저장하는 함수 ---
def save_to_docx(quiz_items, filename="quiz_output.docx"):
    """
    퀴즈 문장 리스트를 받아 새 Word 문서로 저장한다.
    문서 제목을 추가하고 각 퀴즈 문장을 번호와 함께 문단으로 삽입한다.
    저장이 완료되면 사용자에게 저장된 파일명을 출력한다.
    """
    doc = Document()  # 새 문서 객체 생성
    doc.add_heading('O/X 퀴즈 목록', level=1)  # 제목 추가
    for i, item in enumerate(quiz_items, 1):
        doc.add_paragraph(f"{i}. {item}")  # 번호와 문장으로 구성된 문단 추가
    doc.save(filename)  # 파일 저장
    print(f"퀴즈가 Word 파일로 저장되었습니다: {filename}")  # 저장 완료 안내

# --- 메인 프로그램 흐름을 담당하는 함수 ---
def main():
    folder = "study_files"  # 퀴즈 텍스트 파일들이 저장된 폴더명 지정

    while True:  # 사용자 요청에 따라 반복 실행하는 무한 루프
        files = list_files(folder)  # 폴더 내 파일 목록 불러오기
        if not files:  # 파일이 없으면 프로그램 종료 안내 후 종료
            print("프로그램을 종료합니다.")
            break

        try:
            # 사용자로부터 불러올 파일 번호를 입력받음 (1-based 인덱스를 0-based로 변환)
            choice = int(input("\n불러올 파일 번호를 입력하세요: ")) - 1
            # 입력값이 리스트 범위를 벗어나면 경고 메시지 출력 후 다시 입력받기
            if choice < 0 or choice >= len(files):
                print("[경고] 잘못된 번호입니다.\n")
                continue
        except ValueError:
            # 숫자가 아닌 문자를 입력했을 때 발생하는 예외 처리
            print("[경고] 숫자를 입력해야 합니다.\n")
            continue

        filename = files[choice]  # 선택한 파일명
        filepath = os.path.join(folder, filename)  # 파일 경로 생성

        # 지원하지 않는 파일 형식(.txt, .md만 허용)이면 경고 후 다시 입력 받기
        if not filename.endswith(('.txt', '.md')):
            print("[경고] 지원하지 않는 파일 형식입니다.\n")
            continue

        text = load_file(filepath)  # 파일 내용 읽기
        if text:
            quiz_items = generate_ox_quiz(text)  # 문장 분리 후 O/X 붙이기
            print("\nO/X 퀴즈 목록:\n")
            for i, q in enumerate(quiz_items, 1):
                print(f"{i}. {q}")  # 콘솔 출력

            # Word 저장 여부 사용자에게 물어보기
            save_docx = input("\n현재 퀴즈를 Word(.docx) 파일로 저장하시겠습니까? (Y/N): ").strip().lower()
            if save_docx == 'y':
                docx_name = input("저장할 파일 이름을 입력하세요 (확장자 제외): ").strip()
                if not docx_name:  # 이름을 안 입력하면 기본값 사용
                    docx_name = "quiz_output"
                try:
                    # 파일 저장 시도, 실패 시 예외 처리
                    save_to_docx(quiz_items, f"{docx_name}.docx")
                except Exception as e:
                    print(f"[오류] Word 저장에 실패했습니다: {e}")
        else:
            # 파일을 불러오지 못한 경우 안내
            print("[오류] 파일을 불러오지 못했습니다.\n")
            continue  # 다시 입력 받기 위해 루프 처음으로 돌아감

        # 프로그램 재실행 여부 물어보기
        again = input("\n프로그램을 다시 실행하시겠습니까? (Y/N): ").strip().lower()
        if again != 'y':  # y가 아니면 프로그램 종료
            print("프로그램을 종료합니다. 감사합니다.")
            break

# --- 이 파일이 직접 실행될 때 main 함수 실행 ---
if __name__ == "__main__":
    main()
