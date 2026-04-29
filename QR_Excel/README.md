# QR - 엑셀 VBA 모듈 사용 가이드

## 개요

이 프로젝트는 엑셀(VBA)에서 **QR 코드 생성**, **텍스트/JSON 변환**, **JSON 파싱** 기능을 제공하는 세 가지 모듈로 구성됩니다.

| 파일 | 역할 |
|---|---|
| `qr_image.vb` | 셀에 QR 코드 이미지 생성 |
| `text.vb` | 키-코드 매핑, 텍스트 결합, JSON 변환 |
| `JsonConverter.bas` | JSON 파싱 및 직렬화 라이브러리 (VBA-JSON v2.3.1) |

---

## 설치 방법

1. 엑셀 파일을 열고 `Alt + F11`로 VBA 편집기를 엽니다.
2. **모듈 가져오기**: `파일 > 파일 가져오기`에서 세 파일을 순서대로 가져옵니다.
   - `JsonConverter.bas` 먼저 가져와야 합니다 (`ConvertToJson`, `ParseJson` 함수 제공).
   - `text.vb`, `qr_image.vb`는 이후 가져옵니다.
3. `도구 > 참조`에서 **Microsoft Scripting Runtime**이 체크되어 있는지 확인합니다.

> **주의:** `JsonConverter.bas`를 먼저 설치하지 않으면 `text.vb`의 `MyTextJson` 함수에서 오류가 발생합니다.

---

## 모듈별 사용 방법

---

### 1. `qr_image.vb` — QR 코드 이미지 생성

#### `QRCodeImage(text, [level], [version])`

셀에 QR 코드 이미지를 생성합니다. **반드시 셀 수식으로 호출해야 합니다.**

| 매개변수 | 타입 | 필수 | 설명 |
|---|---|---|---|
| `text` | String | ✅ | QR 코드로 인코딩할 텍스트 |
| `level` | String | ❌ | 오류 수정 수준: `L`(저) / `M`(중) / `Q`(고) / `H`(최고), 기본값 `L` |
| `version` | Integer | ❌ | QR 버전 1~40, 마이크로 QR -3~-1, 기본값 `1` |

**셀 수식 예시:**

```
=QRCodeImage(A1)
=QRCodeImage(A1, "M")
=QRCodeImage("https://example.com", "H", 3)
```

#### 주의사항

- **셀에서만 호출 가능**: VBA 코드나 즉시 창에서 직접 실행하면 오류가 발생합니다.
- **임시 파일 생성**: `%TEMP%\qr_셀주소.bmp` 경로에 BMP 파일을 생성합니다. 엑셀 사용 중에는 파일이 잠길 수 있습니다.
- **텍스트가 같으면 재생성하지 않음**: 동일 텍스트를 입력하면 기존 QR 코드를 유지합니다 (성능 최적화).
- **셀 크기에 맞게 이미지 생성**: QR 코드 이미지는 셀(또는 병합 셀)의 크기에 맞게 배치됩니다. 셀이 너무 작으면 QR 코드가 인식되지 않을 수 있습니다.
- **텍스트 길이 제한**: 버전과 오류 수정 수준에 따라 저장 가능한 텍스트 길이가 다릅니다. 너무 긴 텍스트는 `Message too long` 오류가 발생합니다.
- **한글 포함 텍스트**: 한글은 바이너리 모드(UTF-8)로 자동 인코딩됩니다. QR 리더가 UTF-8을 지원하는지 확인하세요.
- **한자(Kanji) 모드**: 한자 모드를 사용하려면 해당 시트의 사용자 정의 속성에 `kanji` 변환 테이블을 설정해야 합니다 (고급 사용자용).

---

### 2. `text.vb` — 텍스트 처리 및 JSON 변환

#### 2-1. `InitKeyCodeDict(keyRange, codeRange)` — 키-코드 딕셔너리 초기화

엑셀 시트의 범위에서 키 이름과 코드 간의 매핑 테이블을 불러옵니다.  
`MyTextJson` 사용 전에 반드시 이 함수를 먼저 호출해야 합니다.

| 매개변수 | 설명 |
|---|---|
| `keyRange` | 키 이름이 있는 셀 범위 (예: `A1:A10`) |
| `codeRange` | 코드가 있는 셀 범위 (예: `B1:B11`) |

- `codeRange`가 `keyRange`보다 1칸 더 크면, 마지막 코드가 **디폴트 코드**로 사용됩니다.
- 디폴트 코드를 지정하지 않으면 `"mr"`이 기본값입니다.

**VBA 코드 예시:**

```vb
Call InitKeyCodeDict(Sheet1.Range("A1:A5"), Sheet1.Range("B1:B6"))
```

---

#### 2-2. `GetKeyCode(keyName)` — 키에 대응하는 코드 반환

| 매개변수 | 설명 |
|---|---|
| `keyName` | 조회할 키 이름 |

- 딕셔너리에 키가 있으면 대응 코드를 반환합니다.
- 키가 없으면 **원래 키 이름을 그대로 반환**합니다 (오류 없음).
- `InitKeyCodeDict`를 호출하지 않은 상태에서 사용하면 키 이름을 그대로 반환합니다.

---

#### 2-3. `MyTextJoin(Delimiter, IgnoreEmpty, TargetRange...)` — 텍스트 결합

엑셀 내장 `TEXTJOIN` 함수와 동일한 역할을 합니다.

| 매개변수 | 설명 |
|---|---|
| `Delimiter` | 구분자 문자열 (예: `","`, `"-"`) |
| `IgnoreEmpty` | 빈 셀 무시 여부 (`True` / `False`) |
| `TargetRange...` | 결합할 셀 범위 (여러 범위 지정 가능) |

**셀 수식 예시:**

```
=MyTextJoin(",", True, A1:A10)
=MyTextJoin("-", False, A1:A5, C1:C5)
```

#### 주의사항

- 엑셀 2016 이상에서는 내장 `TEXTJOIN`을 사용하는 것을 권장합니다.
- 빈 셀을 포함하고 싶으면 `IgnoreEmpty`를 `False`로 설정하세요.

---

#### 2-4. `MyTextJson(rootKey, args...)` — JSON 문자열 생성

키-값 쌍을 JSON 형식의 문자열로 변환합니다.

| 매개변수 | 설명 |
|---|---|
| `rootKey` | 최상위 JSON 키 이름. 빈 문자열(`""`)이면 루트 없이 생성 |
| `args...` | 키-값 쌍 순서로 전달 (셀 범위 또는 직접 값 모두 가능) |

**셀 수식 예시:**

```
' 결과: {"order":{"name":"홍길동","qty":"5"}}
=MyTextJson("order", "name", A1, "qty", B1)

' 결과: {"name":"홍길동","qty":"5"}
=MyTextJson("", "name", A1, "qty", B1)

' 홀수 개 입력 시 마지막 항목은 디폴트 코드("mr")를 키로 사용
' 결과: {"order":{"name":"홍길동","mr":"메모"}}
=MyTextJson("order", "name", A1, C1)
```

#### 주의사항

- **`InitKeyCodeDict` 선행 필요**: 키 이름을 코드로 자동 변환하는 기능을 쓰려면 `InitKeyCodeDict`를 먼저 실행해야 합니다.
- **`args`는 반드시 짝수 개 권장**: 홀수 개이면 마지막 값의 키가 디폴트 코드(`"mr"`)로 지정됩니다.
- **중복 키 불가**: `args` 내에 같은 키가 두 번 등장하면 VBA 런타임 오류가 발생합니다.
- **셀 범위 전달 시**: 범위 내 모든 셀 값이 순서대로 펼쳐져 키-값 쌍으로 처리됩니다.

---

### 3. `JsonConverter.bas` — JSON 라이브러리 (VBA-JSON v2.3.1)

Tim Hall의 오픈소스 라이브러리입니다. 직접 수정하지 마세요.

#### 주요 공개 함수

| 함수 | 설명 |
|---|---|
| `ParseJson(jsonString)` | JSON 문자열을 Dictionary/Collection 객체로 변환 |
| `ConvertToJson(value, [whitespace])` | Dictionary/Collection/배열을 JSON 문자열로 변환 |
| `ParseUtc(date)` | UTC 날짜를 로컬 날짜로 변환 |
| `ConvertToUtc(date)` | 로컬 날짜를 UTC 날짜로 변환 |
| `ParseIso(isoString)` | ISO 8601 문자열을 날짜로 변환 |
| `ConvertToIso(date)` | 날짜를 ISO 8601 문자열로 변환 |

#### 주의사항

- **15자리 초과 숫자**: VBA의 숫자 정밀도 한계(15자리)로 인해 긴 정수(예: 카드번호, ID)는 자동으로 문자열로 처리됩니다. 강제로 `Double`을 사용하려면 `JsonOptions.UseDoubleForLargeNumbers = True`로 설정하세요.
- **키 따옴표**: JSON 표준에 따라 키에 반드시 큰따옴표가 필요합니다. 따옴표 없는 키를 허용하려면 `JsonOptions.AllowUnquotedKeys = True`로 설정하세요.
- **슬래시 이스케이프**: `/` 문자는 기본적으로 이스케이프하지 않습니다. 필요하면 `JsonOptions.EscapeSolidus = True`로 설정하세요.

---

## 함수 간 의존 관계

```
JsonConverter.bas
    └── ConvertToJson()
            └── text.vb > MyTextJson()
                            └── GetKeyCode()
                                    └── InitKeyCodeDict() (선행 호출 필요)
```

---

## 자주 발생하는 오류

| 오류 메시지 | 원인 | 해결 방법 |
|---|---|---|
| `Call only from sheet` | `QRCodeImage`를 VBA 코드에서 직접 호출함 | 반드시 셀 수식으로 호출 |
| `Message too long` | QR 텍스트가 버전/수준 한계 초과 | `version` 값을 높이거나 텍스트를 줄임 |
| `중복 키` 런타임 오류 | `MyTextJson`에 같은 키 두 번 전달 | `args` 내 키 중복 제거 |
| `Error parsing JSON: 10001` | `ParseJson`에 잘못된 JSON 전달 | JSON 형식(따옴표, 괄호 등) 확인 |
| `JsonConverter` 없음 오류 | `JsonConverter.bas` 미설치 | VBA 편집기에서 `JsonConverter.bas` 먼저 가져오기 |

---

## 라이선스

- `JsonConverter.bas`: MIT License — © Tim Hall (https://github.com/VBA-tools/VBA-JSON)
- `qr_image.vb`: ISO/IEC 18004:2006 기반 구현체
- `text.vb`: 프로젝트 내부 모듈
