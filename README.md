# windows_exporter 인스톨러

<p align="center">
  <a href="LICENSE">
    <img src=https://img.shields.io/badge/License-MIT-lightgrey.svg?longCache=true" alt="MIT License">
  </a>
  <a href="https://python.org/downloads">
    <img alt="Python" src="https://img.shields.io/badge/Python-3776AB.svg?style=flat-square&logo=Python&logoColor=white">
  </a>
  <a href="https://prometheus.io/download">
    <img alt="Prometheus" src="https://img.shields.io/badge/Prometheus-E6522C?styel=flat-square&logo=Prometheus&logoColor=white" >
  </a>
</p>

## Introduction

> 통합모니터링 대시보드 구축 프로젝트 중 윈도우 서버들을 대상으로 Prometheus의 windows_exporter의 일관적인 deployment를 위해 다운로드, 윈도우 서비스 등록/제거 등의 작업을 표준화하도록 하는 유틸리티

이 유틸리티 프로그램은 다음과 같은 이유로 제작함
- 장기간에 걸쳐 다수의 윈도우서버에 적용, 운영하는 경우 exporter를 설치하는 시기마다 최신 버전을 적용하면 버전 관리가 안되며 버그 발생 시 추적이 어려워 버전을 특정하여 고정 설치를 도움
- 단일 실행 파일만 다운로드하여 단순 구동하는 경우 시스템 재시작 또는 알수 없는 이유로 프로세스가 중단 된 경우 수동으로 시작해주기 전까지 시스템 메트릭 수집이 불가하므로 윈도우 서비스에 등록 시켜 exporter 구동 관리
- 작업 순서를 UI로 제공하여 휴먼 에러를 예방
- 서비스 등록 실패시 원복을 할 수 있도록 별도 탭에 서비스 제거 기능

## Features

- 서비스 제어(등록/제거)를 위한 관리자 권한 실행
- windows_exporter 공식 Github 경로 제공
- Github Repository의 releases 페이지에서 설치파일(msi) 직접 다운로드 및 설치
- msi 설치가 불가능한 윈도우 서버에서 exe 파일을 수동으로 설치, 등록하도록 지원
- 설치 확인 페이지 (localhost:8192/metrics) 링크
- 윈도우 서비스 등록 (이름, 설명) 제공

## Install library & Compile

```bash
#pip install pywin32 requests
pip install -r requirements.txt

pip install pyinstaller
pyinstaller -F --noconsole --add-data "github_icon.png;." --add-data "logo.png;." wexporter-installer.py
```

## Releases Notes

### 0.3.3

- '서비스 제거' 탭에 '서비스 열기' 버튼 생성 및 기존 버튼 정렬

### 0.3.2

- Windows Resizeable 옵션 변경 (False)

### 0.3.1

- '서비스 설치' 버튼 재배치
- 아이콘 변경 (github, web)
- title_frame 여백 조정
- prometheus 설정 가이드 텍스트 고정폭글꼴 적용 (Consolas)

### 0.3.0

- 코드 전체 Refactoring 작업
- 함수이름 snake case로 통일
- 예외처리 추가 (다운로드 오류, 설치 오류, 예상치 못한 오류)
