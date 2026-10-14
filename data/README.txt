가나다생명 심사부 AI PoC 더미 데이터 (전부 가상 데이터)

[documents/]  신구 문서 3쌍(6개) + 분류 불가 테스트용 회의록 1개
[documents.csv]  문서 메타 (doc_id, 유형, 버전, 파일명 등)
[expected_changes.csv]  정답표: 앱이 찾아내야 하는 변경 15건 (개정 9 / 신설 3 / 삭제 3)
[users.csv]  알림 대상 담당자 5명과 구독 유형 (구분자 |)
[classification_keywords.csv]  문서유형 분류 키워드
[mail_samples/]  알림 메일 본문 샘플
[changes_template.xlsx]  변경대비표 양식 (조항|현행|개정|변경유형|비고)

검증 방법: 앱에 v1.0/v2.0 파일을 올려 나온 결과를 expected_changes.csv와 비교하세요.
참고: 신설 조항은 번호 밀림을 막기 위해 '제N조의2' 형태로 되어 있습니다.
