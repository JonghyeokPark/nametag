# Nametag

학회 이름표 생성 자동화 툴

## Build
```
git clone https://github.com/JonghyeokPark/nametag

python3 -m venv .venv
source .venv/bin/activate
```

## Install Python Packages

```
pip install openpyxl
pip install python-pptx
```

## Prepare

학회 이름표를 생성하기 위해서는 다음과 같은 파일을 준비해야 합니다.   
- `template.pptx` : 학회 이름표 형식 템플릿입니다. 본 파일은 [NVRAMOS 2024](https://sigfast.or.kr/nvramos/nvramos24/) 학회 이름표 템플릿 파일입니다.
- `list.xlsx` : 학회 참석 등록자 명단
   
`template.pptx` 파일은 파워포인트 개체 요소 (예: 텍스트 상자)가 학회 참석등록자 이름, 소속에 맞게 미리 개체이름이 지정이 되었습니다.
파워포인트에서 `편집 > 선택 > 선택창` 을 누르면 해당 개체 이름을 지정할 수 있습니다.
본 프로젝트에서는 `name` 과 `affiliation` 으로 지정하였고, 한 slide 당 4개의 이름표를 생성합니다.


## Run

```
python3 run.py
```

## Contact

- Author: Jonghyeok Park 
- Homepage: [IDS Lab.](http://ids.hufs.ac.kr)
- E-mail: jonghyeok_park@korea.ac.kr
