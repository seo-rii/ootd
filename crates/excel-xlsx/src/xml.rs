use office_common::{OmError, OmResult};
use quick_xml::events::BytesRef;
use quick_xml::name::{LocalName, NamespaceResolver, QName, ResolveResult};

pub(crate) fn decode_general_reference(
    reference: &BytesRef<'_>,
    context: &str,
) -> OmResult<String> {
    let reference = reference
        .decode()
        .map_err(|error| OmError::parse(error.to_string()))?;
    let value = if let Some(number) = reference.strip_prefix("#x") {
        let codepoint = u32::from_str_radix(number, 16).map_err(|_| {
            OmError::parse(format!(
                "{context}: invalid XML character reference: &{reference};"
            ))
        })?;
        char::from_u32(codepoint)
            .ok_or_else(|| {
                OmError::parse(format!(
                    "{context}: invalid XML character reference: &{reference};"
                ))
            })?
            .to_string()
    } else if let Some(number) = reference.strip_prefix("#X") {
        let codepoint = u32::from_str_radix(number, 16).map_err(|_| {
            OmError::parse(format!(
                "{context}: invalid XML character reference: &{reference};"
            ))
        })?;
        char::from_u32(codepoint)
            .ok_or_else(|| {
                OmError::parse(format!(
                    "{context}: invalid XML character reference: &{reference};"
                ))
            })?
            .to_string()
    } else if let Some(number) = reference.strip_prefix('#') {
        let codepoint = number.parse::<u32>().map_err(|_| {
            OmError::parse(format!(
                "{context}: invalid XML character reference: &{reference};"
            ))
        })?;
        char::from_u32(codepoint)
            .ok_or_else(|| {
                OmError::parse(format!(
                    "{context}: invalid XML character reference: &{reference};"
                ))
            })?
            .to_string()
    } else {
        match reference.as_ref() {
            "amp" => "&",
            "lt" => "<",
            "gt" => ">",
            "quot" => "\"",
            "apos" => "'",
            _ => {
                return Err(OmError::parse(format!(
                    "{context}: unknown XML entity reference: &{reference};"
                )));
            }
        }
        .to_string()
    };
    Ok(value)
}

pub(crate) fn resolved_element_is(
    namespace: &ResolveResult<'_>,
    local_name: LocalName<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    local_name.as_ref() == expected_local_name
        && matches!(
            namespace,
            ResolveResult::Bound(namespace) if namespace.as_ref() == expected_namespace
        )
}

pub(crate) fn expanded_name_is(
    namespace: Option<&[u8]>,
    local_name: LocalName<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    namespace == Some(expected_namespace) && local_name.as_ref() == expected_local_name
}

pub(crate) fn unqualified_attribute_is(
    resolver: &NamespaceResolver,
    name: QName<'_>,
    expected_local_name: &[u8],
) -> bool {
    let (namespace, local_name) = resolver.resolve_attribute(name);
    local_name.as_ref() == expected_local_name && matches!(namespace, ResolveResult::Unbound)
}

pub(crate) fn namespaced_attribute_is(
    resolver: &NamespaceResolver,
    name: QName<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    let (namespace, local_name) = resolver.resolve_attribute(name);
    local_name.as_ref() == expected_local_name
        && matches!(
            namespace,
            ResolveResult::Bound(namespace) if namespace.as_ref() == expected_namespace
        )
}

pub(crate) fn qualified_name_like(reference: &[u8], local_name: &str) -> String {
    let Some(separator) = reference.iter().position(|byte| *byte == b':') else {
        return local_name.to_string();
    };
    format!(
        "{}:{local_name}",
        String::from_utf8_lossy(&reference[..separator])
    )
}
