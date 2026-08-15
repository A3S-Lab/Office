use std::collections::{BTreeMap, HashSet};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::updates::decoder::Decode;
use yrs::{Any, Array, Doc, Map, MapRef, Out, ReadTxn, Transact, Update};

const PDF_COLLABORATION_FIXTURE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3C2ZpeHR1cmUtcGRmKAETYTNzLm9mZmljZS5tZXRhZGF0YQRraW5kAXcDcGRmIQETYTNzLm9mZmljZS5tZXRhZGF0YQtpbml0aWFsaXplZAEIASFhM3Mub2ZmaWNlLmJvb3RzdHJhcC5pbml0aWFsaXplcnMBdxA0MjQyNDI6YW5vbnltb3VzKAEVYTNzLm9mZmljZS5wZGYuc291cmNlBnNoYTI1NgF3QDAxMjM0NTY3ODlhYmNkZWYwMTIzNDU2Nzg5YWJjZGVmMDEyMzQ1Njc4OWFiY2RlZjAxMjM0NTY3ODlhYmNkZWYoARVhM3Mub2ZmaWNlLnBkZi5zb3VyY2UKYnl0ZUxlbmd0aAF9gEAoARVhM3Mub2ZmaWNlLnBkZi5zb3VyY2UJcGFnZUNvdW50AX0DCAEgYTNzLm9mZmljZS5wZGYuc291cmNlLWlkZW50aXRpZXMBd217ImJ5dGVMZW5ndGgiOjQwOTYsInBhZ2VDb3VudCI6Mywic2hhMjU2IjoiMDEyMzQ1Njc4OWFiY2RlZjAxMjM0NTY3ODlhYmNkZWYwMTIzNDU2Nzg5YWJjZGVmMDEyMzQ1Njc4OWFiY2RlZiJ9KAEjYTNzLm9mZmljZS5wZGYuZm9ybS12YWx1ZXMucHJlc2VuY2UOQXBwbGljYW50Lk5hbWUBeCgBIWEzcy5vZmZpY2UucGRmLmZvcm0tdmFsdWVzLmZpZWxkcydbIkFwcGxpY2FudC5OYW1lIiwiW1widmFsdWVcIixcImlkXCJdIl0Bdw5BcHBsaWNhbnQuTmFtZSgBIWEzcy5vZmZpY2UucGRmLmZvcm0tdmFsdWVzLmZpZWxkcypbIkFwcGxpY2FudC5OYW1lIiwiW1widmFsdWVcIixcInZhbHVlXCJdIl0BdwNBZGEIASBhM3Mub2ZmaWNlLnBkZi5mb3JtLXZhbHVlcy5vcmRlcgF3DkFwcGxpY2FudC5OYW1lqLLyGQQBeAGy8hkBBAE=";

#[derive(Debug, PartialEq, Eq)]
pub struct PdfSnapshot {
    pub annotations: BTreeMap<String, serde_json::Value>,
    pub form_values: BTreeMap<String, String>,
    pub form_order: Vec<String>,
    pub redaction_proposals: BTreeMap<String, serde_json::Value>,
    pub page_operations: BTreeMap<String, serde_json::Value>,
    pub review_decisions: BTreeMap<String, serde_json::Value>,
    pub record_claims: Vec<String>,
}

pub fn pdf_collaboration_fixture() -> Vec<u8> {
    STANDARD.decode(PDF_COLLABORATION_FIXTURE_BASE64).unwrap()
}

pub fn pdf_snapshot(update: &[u8]) -> PdfSnapshot {
    let document = Doc::with_client_id(710_013);
    document
        .transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    let presence = document.get_or_insert_map("a3s.office.pdf.form-values.presence");
    let fields = document.get_or_insert_map("a3s.office.pdf.form-values.fields");
    let order = document.get_or_insert_array("a3s.office.pdf.form-values.order");
    let annotation_presence = document.get_or_insert_map("a3s.office.pdf.annotations.presence");
    let annotation_fields = document.get_or_insert_map("a3s.office.pdf.annotations.fields");
    let redaction_presence =
        document.get_or_insert_map("a3s.office.pdf.redaction-proposals.presence");
    let redaction_fields = document.get_or_insert_map("a3s.office.pdf.redaction-proposals.fields");
    let page_operation_presence =
        document.get_or_insert_map("a3s.office.pdf.page-operations.presence");
    let page_operation_fields = document.get_or_insert_map("a3s.office.pdf.page-operations.fields");
    let decision_presence = document.get_or_insert_map("a3s.office.pdf.review-decisions.presence");
    let decision_fields = document.get_or_insert_map("a3s.office.pdf.review-decisions.fields");
    let claims = document.get_or_insert_array("a3s.office.pdf.record-claims");
    let transaction = document.transact();
    let form_values = presence
        .iter(&transaction)
        .map(|(field_id, present)| {
            assert!(matches!(present, Out::Any(Any::Bool(true))));
            let key = form_field_key(field_id, "value");
            let value = match fields.get(&transaction, key.as_str()) {
                Some(Out::Any(Any::String(value))) => value.to_string(),
                value => panic!("unexpected PDF form value: {value:?}"),
            };
            (field_id.to_owned(), value)
        })
        .collect();
    let form_order = (0..order.len(&transaction))
        .map(|index| match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            value => panic!("unexpected PDF form order entry: {value:?}"),
        })
        .collect();
    let annotations = record_snapshot(&annotation_presence, &annotation_fields, &transaction);
    let redaction_proposals = record_snapshot(&redaction_presence, &redaction_fields, &transaction);
    let page_operations = record_snapshot(
        &page_operation_presence,
        &page_operation_fields,
        &transaction,
    );
    let review_decisions = record_snapshot(&decision_presence, &decision_fields, &transaction);
    let record_claims = (0..claims.len(&transaction))
        .map(|index| match claims.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            value => panic!("unexpected PDF record claim: {value:?}"),
        })
        .collect();
    PdfSnapshot {
        annotations,
        form_values,
        form_order,
        redaction_proposals,
        page_operations,
        review_decisions,
        record_claims,
    }
}

fn record_snapshot<T: ReadTxn>(
    presence: &MapRef,
    fields: &MapRef,
    transaction: &T,
) -> BTreeMap<String, serde_json::Value> {
    let records = presence
        .iter(transaction)
        .map(|(id, present)| {
            assert!(matches!(present, Out::Any(Any::Bool(true))));
            (id.to_owned(), Vec::new())
        })
        .collect::<BTreeMap<_, _>>();
    let mut records = records;
    for (encoded, value) in fields.iter(transaction) {
        let identity = serde_json::from_str::<Vec<String>>(encoded).unwrap();
        assert_eq!(identity.len(), 2);
        let Some(record) = records.get_mut(&identity[0]) else {
            continue;
        };
        let field = serde_json::from_str::<Vec<String>>(&identity[1]).unwrap();
        assert!(field.len() >= 2);
        let kind = field[0].clone();
        let value = match (kind.as_str(), value) {
            ("object", Out::Any(Any::Bool(true))) => serde_json::Value::Bool(true),
            ("value", Out::Any(value)) => serde_json::to_value(value).unwrap(),
            _ => panic!("unexpected shared PDF record field"),
        };
        record.push((kind, field[1..].to_vec(), value));
    }
    records
        .into_iter()
        .map(|(id, mut entries)| {
            entries.sort_by(|left, right| {
                left.1.len().cmp(&right.1.len()).then_with(|| {
                    if left.0 == right.0 {
                        std::cmp::Ordering::Equal
                    } else if left.0 == "object" {
                        std::cmp::Ordering::Less
                    } else {
                        std::cmp::Ordering::Greater
                    }
                })
            });
            let object_paths = entries
                .iter()
                .filter(|(kind, _, _)| kind == "object")
                .map(|(_, path, _)| path.clone())
                .collect::<HashSet<_>>();
            for (_, path, _) in &entries {
                for depth in 1..path.len() {
                    assert!(object_paths.contains(&path[..depth]));
                }
            }
            let mut record = serde_json::Map::new();
            for (kind, path, value) in entries {
                if kind == "object" {
                    ensure_json_object_path(&mut record, &path);
                } else {
                    set_json_value(&mut record, &path, value);
                }
            }
            (id, serde_json::Value::Object(record))
        })
        .collect()
}

fn ensure_json_object_path(
    object: &mut serde_json::Map<String, serde_json::Value>,
    path: &[String],
) {
    let (key, rest) = path.split_first().expect("non-empty PDF field path");
    let child = object
        .entry(key.clone())
        .or_insert_with(|| serde_json::Value::Object(serde_json::Map::new()))
        .as_object_mut()
        .expect("PDF object marker must not overlap a value");
    if !rest.is_empty() {
        ensure_json_object_path(child, rest);
    }
}

fn set_json_value(
    object: &mut serde_json::Map<String, serde_json::Value>,
    path: &[String],
    value: serde_json::Value,
) {
    let (key, rest) = path.split_first().expect("non-empty PDF field path");
    if rest.is_empty() {
        assert!(object.insert(key.clone(), value).is_none());
        return;
    }
    let child = object
        .get_mut(key)
        .and_then(serde_json::Value::as_object_mut)
        .expect("PDF value must have an object marker parent");
    set_json_value(child, rest, value);
}

fn form_field_key(field_id: &str, property: &str) -> String {
    let field = serde_json::to_string(&["value", property]).unwrap();
    serde_json::to_string(&[field_id, field.as_str()]).unwrap()
}
