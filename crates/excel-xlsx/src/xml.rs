use quick_xml::name::{LocalName, NamespaceResolver, QName, ResolveResult};

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
