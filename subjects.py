# -*- coding: utf-8 -*-
"""과목별 설정. Railway 환경변수 SUBJECT(math/english/korean)로 선택한다.

과목마다 다른 것:
  - display : 안내문·보고서에 쓰는 과목 표시 이름
  - schedules : {반: [수업요일idx,...]}  (비우면 출석부 파일 날짜에서 자동 추론)
  - times : {반: {요일idx(str): 'HH:MM'}}  출석 미입력 알림 시각 (비우면 알림 명령으로 설정)

수강생 명단은 출석부 엑셀(시드)에서, 담당 선생님은 '담당' 명령으로 지정하므로
여기엔 넣지 않는다. 새 과목은 이 파일에 항목만 추가하면 된다.
"""
import attendance_core as ac

SUBJECTS = {
    # 수학 — 기존 값 그대로 사용
    "math": {
        "display": "Zest 수학과",
        "schedules": ac.DEFAULT_SCHEDULES,
        "times": ac.DEFAULT_CLASS_TIMES,
    },
    # 영어 — 분반·시간은 배포 후 '일정'/'알림' 명령 또는 아래에 채워 넣기
    "english": {
        "display": "Zest 영어과",
        "schedules": {},
        "times": {},
    },
    # 국어
    "korean": {
        "display": "Zest 국어과",
        "schedules": {},
        "times": {},
    },
}


def get(name):
    """과목 설정 반환. 모르는 이름이면 math."""
    return SUBJECTS.get((name or "math").strip().lower(), SUBJECTS["math"])
