"""오류 코드. VDI 화면 캡처 한 장만 보고 원인과 조치를 알 수 있게 한다."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ErrorCode:
    code: str
    title: str
    action: str


_CODES = [
    ErrorCode("E100", "pywin32 를 불러올 수 없음", "setup.bat 을 먼저 실행하세요."),
    ErrorCode("E101", "Outlook 연결 실패", "Outlook 을 실행해 로그인된 상태에서 다시 시도하세요."),
    ErrorCode("E102", "폴더 목록 읽기 실패", "Outlook 이 응답하는지 확인하고 다시 시도하세요."),
    ErrorCode("E103", "접근할 수 없는 폴더가 있음", "해당 폴더는 수집에서 빠집니다. 공유/권한 폴더인지 확인하세요."),
    ErrorCode("E104", "폴더의 목록(Table) 읽기 실패", "해당 폴더는 수집에서 빠질 수 있습니다. 이 화면을 전달해 주세요."),
    ErrorCode("E105", "수신 시각 정렬이 기대와 다름", "증분 수집을 쓸 수 없습니다. 이 화면을 전달해 주세요."),
    ErrorCode("E106", "Message-ID 가 없는 항목이 많음", "아래 키 비교 결과로 대체 키를 정합니다. 동작에는 문제 없습니다."),
    ErrorCode("E107", "아카이브 바로가기(stub) 항목 발견", "본문이 잘려 있을 수 있습니다. 원본 확보 방법을 별도로 검증합니다."),
    ErrorCode("E108", "증분 조회 결과 불일치", "수집 누락 위험이 있습니다. 이 화면을 전달해 주세요."),
    ErrorCode("E109", ".msg 저장 실패", "보존 경로의 쓰기 권한과 남은 용량을 확인하세요."),
    ErrorCode("E110", "보존 경로에 쓸 수 없음", "다른 경로를 지정하세요. (예: 개인 네트워크 드라이브)"),
    ErrorCode("E111", "Windows 긴 경로(260자 초과) 미지원", "보존 경로를 짧게 지정하세요. (예: H:\\mail)"),
    ErrorCode("E112", "저장소의 폴더 읽기 실패", "해당 저장소는 수집에서 빠집니다. 이 화면을 전달해 주세요."),
    ErrorCode("E113", "메일 열기(상세 읽기) 실패", "본문 수집이 안 될 수 있습니다. 이 화면을 전달해 주세요."),
]

CODES = {c.code: c for c in _CODES}
