# 관측소 카탈로그(오프라인 후보 탐색용) — 스펙

> 목적: 외부 API가 일시 장애여도 “가까운 관측소 후보 3개”를 안정적으로 선정하기 위해, 최소한의 관측소 좌표 카탈로그를 로컬 파일로 유지한다.

## 파일 목록(권장)
- `kma_asos_stations.csv`
- `airkorea_stations.csv`

## CSV 컬럼(권장; 최소)
- `station_id` (string)
- `station_name` (string)
- `lat` (float, EPSG:4326)
- `lon` (float, EPSG:4326)
- `admin_si` (string, optional)
- `admin_sigungu` (string, optional)

## 사용 위치(계획)
- Planner가 `LOCATION.center_lat/center_lon` 기준으로 거리 계산 → 후보 3개를 `DATA_REQUESTS.params_json.station_candidates`에 채움.  
  refs: `eia-gen/docs/13_data_requests_and_connectors_spec.md:73`


## 생성/갱신(권장)
- `eia-gen fetch-kma-asos-stations` → `config/stations/kma_asos_stations.csv` 생성
