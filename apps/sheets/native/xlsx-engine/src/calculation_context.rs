//! Host-only calculation resources. Hidden definition bodies must never be
//! added to public open metadata or forwarded through a renderer bridge.
use super::*;
use std::collections::BTreeMap;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculationName {
    pub name: String,
    pub formula: String,
    pub hidden: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculationContext {
    pub session_id: String,
    pub names: Vec<CalculationName>,
    /// Preserve all calcPr attributes, including iteration and precision. The
    /// host must interpret these rather than silently assuming defaults.
    pub calculation_properties: BTreeMap<String, String>,
    /// Unresolved dependencies are explicit; their presence is not permission
    /// to fetch external files or claim a fully hydrated engine snapshot.
    pub external_relationship_ids: Vec<String>,
}

impl WorkbookSessions {
    pub fn read_calculation_context(
        &self,
        session_id: &str,
        cancelled: &AtomicBool,
    ) -> Result<CalculationContext, SidecarError> {
        if cancelled.load(Ordering::Acquire) {
            return Err(SidecarError::cancelled());
        }
        let session = self
            .sessions
            .get(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let mut archive = ZipArchive::new(File::open(&session.path)?)?;
        let xml = read_zip_string(&mut archive, "xl/workbook.xml")?;
        let mut reader = quick_xml::NsReader::from_str(&xml);
        let mut result = CalculationContext {
            session_id: session_id.into(),
            names: Vec::new(),
            calculation_properties: BTreeMap::new(),
            external_relationship_ids: Vec::new(),
        };
        let mut current: Option<CalculationName> = None;
        let mut path: Vec<Option<Vec<u8>>> = Vec::new();
        let mut saw_root = false;
        loop {
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::cancelled());
            }
            let (namespace, event) = reader.read_resolved_event()?;
            let main = matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if
                ns.as_ref() == b"http://schemas.openxmlformats.org/spreadsheetml/2006/main" ||
                ns.as_ref() == b"http://purl.oclc.org/ooxml/spreadsheetml/main");
            match event {
                event @ (Event::Start(_) | Event::Empty(_)) => {
                    let empty = matches!(&event, Event::Empty(_));
                    let element = match event {
                        Event::Start(e) | Event::Empty(e) => e,
                        _ => unreachable!(),
                    };
                    let local = element.local_name();
                    if path.is_empty() {
                        if saw_root || !main || local.as_ref() != b"workbook" {
                            return Err(SidecarError::Workbook(
                                "Invalid calculation workbook root.".into(),
                            ));
                        }
                        saw_root = true;
                    }
                    if current.is_some() {
                        return Err(SidecarError::Workbook(
                            "Nested markup in a defined name.".into(),
                        ));
                    }
                    let parent = |expected: &[&[u8]]| {
                        path.len() == expected.len()
                            && path
                                .iter()
                                .zip(expected)
                                .all(|(actual, expected)| actual.as_deref() == Some(*expected))
                    };
                    if main && local.as_ref() == b"calcPr" && parent(&[b"workbook"]) {
                        for attribute in element.attributes() {
                            let attribute = attribute
                                .map_err(|error| SidecarError::Workbook(error.to_string()))?;
                            if !attribute.key.as_ref().contains(&b':')
                                && attribute.key.as_ref() != b"xmlns"
                            {
                                result.calculation_properties.insert(
                                    String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                                    attribute
                                        .decode_and_unescape_value(reader.decoder())?
                                        .into_owned(),
                                );
                            }
                        }
                    }
                    if main
                        && local.as_ref() == b"externalReference"
                        && parent(&[b"workbook", b"externalReferences"])
                    {
                        let mut id = None;
                        for attribute in element.attributes() {
                            let attribute = attribute
                                .map_err(|error| SidecarError::Workbook(error.to_string()))?;
                            let (namespace, local) = reader.resolve_attribute(attribute.key);
                            if local.as_ref() == b"id"
                                && matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if
                                ns.as_ref() == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships" ||
                                ns.as_ref() == b"http://purl.oclc.org/ooxml/officeDocument/relationships")
                            {
                                id = Some(
                                    attribute
                                        .decode_and_unescape_value(reader.decoder())?
                                        .into_owned(),
                                );
                            }
                        }
                        result.external_relationship_ids.push(
                            id.filter(|id| !id.is_empty()).ok_or_else(|| {
                                SidecarError::Workbook(
                                    "External reference is missing its relationship id.".into(),
                                )
                            })?,
                        );
                    }
                    if main
                        && local.as_ref() == b"definedName"
                        && parent(&[b"workbook", b"definedNames"])
                    {
                        let name = plain_attribute(&reader, &element, b"name")?
                            .filter(|name| !name.is_empty())
                            .ok_or_else(|| {
                                SidecarError::Workbook("Defined name is missing its name.".into())
                            })?;
                        let sheet_id = match plain_attribute(&reader, &element, b"localSheetId")? {
                            None => None,
                            Some(index) => Some(
                                session
                                    .sheets
                                    .get(index.parse::<usize>().map_err(|_| {
                                        SidecarError::Workbook("Invalid defined-name scope.".into())
                                    })?)
                                    .ok_or_else(|| {
                                        SidecarError::Workbook("Unknown defined-name scope.".into())
                                    })?
                                    .id
                                    .clone(),
                            ),
                        };
                        let entry = CalculationName {
                            name,
                            sheet_id,
                            formula: String::new(),
                            hidden: plain_attribute(&reader, &element, b"hidden")?
                                .is_some_and(|value| value == "1" || value == "true"),
                        };
                        if empty {
                            result.names.push(entry);
                        } else {
                            current = Some(entry);
                        }
                    }
                    if !empty {
                        path.push(main.then(|| local.as_ref().to_vec()));
                    }
                }
                Event::Text(text) => {
                    if let Some(name) = &mut current {
                        name.formula.push_str(&decode_text(&text)?);
                    }
                }
                Event::CData(text) => {
                    if let Some(name) = &mut current {
                        name.formula.push_str(&decode_cdata(&text)?);
                    }
                }
                Event::GeneralRef(reference) => {
                    if let Some(name) = &mut current {
                        name.formula.push_str(&general_ref_text(&reference)?);
                    }
                }
                Event::End(_) => {
                    if let Some(mut name) = current.take() {
                        name.formula = strip_future_function_markers(&name.formula);
                        result.names.push(name);
                    }
                    path.pop();
                }
                Event::DocType(_) => {
                    return Err(SidecarError::Workbook(
                        "DTD calculation context is unsupported.".into(),
                    ));
                }
                Event::Eof => {
                    if current.is_some() || !path.is_empty() || !saw_root {
                        return Err(SidecarError::Workbook(
                            "Incomplete calculation workbook.".into(),
                        ));
                    }
                    break;
                }
                _ => {}
            }
        }
        Ok(result)
    }
}

fn plain_attribute(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    key: &[u8],
) -> Result<Option<String>, SidecarError> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| SidecarError::Workbook(error.to_string()))?;
        if attribute.key.as_ref() == key {
            value = Some(
                attribute
                    .decode_and_unescape_value(reader.decoder())?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tests::open_fixture;

    fn read(body: &str, prefix: bool) -> Result<CalculationContext, SidecarError> {
        let workbook = format!(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:foreign="urn:foreign"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>{body}</workbook>"#
        );
        let workbook = if prefix {
            // Preserve comments/CDATA verbatim while prefixing only element tags.
            workbook
                .replace("xmlns=", "xmlns:x=")
                .replace("<workbook", "<x:workbook")
                .replace("</workbook", "</x:workbook")
                .replace("<sheets", "<x:sheets")
                .replace("</sheets", "</x:sheets")
                .replace("<sheet ", "<x:sheet ")
                .replace("<definedNames", "<x:definedNames")
                .replace("</definedNames", "</x:definedNames")
                .replace("<definedName ", "<x:definedName ")
                .replace("</definedName>", "</x:definedName>")
                .replace("<calcPr", "<x:calcPr")
                .replace("<externalReferences", "<x:externalReferences")
                .replace("</externalReferences", "</x:externalReferences")
                .replace("<externalReference ", "<x:externalReference ")
        } else {
            workbook
        };
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#;
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let (_dir, path) = open_fixture(&[
            ("xl/workbook.xml", &workbook),
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/worksheets/sheet1.xml", sheet),
            ("xl/worksheets/sheet2.xml", sheet),
        ]);
        let mut sessions = WorkbookSessions::new();
        let opened = sessions.open(&path)?;
        let context =
            sessions.read_calculation_context(&opened.session_id, &AtomicBool::new(false))?;
        assert!(matches!(
            sessions.read_calculation_context(&opened.session_id, &AtomicBool::new(true)),
            Err(SidecarError::Cancelled)
        ));
        // Hidden bodies remain absent from ordinary open metadata.
        assert!(
            !opened
                .defined_names
                .iter()
                .any(|name| name.name == "Secret")
        );
        Ok(context)
    }

    #[test]
    fn preserves_private_scoped_empty_and_builtin_names_with_calc_properties() {
        for prefix in [false, true] {
            let context = read(r#"<definedNames><definedName name="Rate">Data!$A$1</definedName><definedName name="Rate" localSheetId="1">Other!$B$1</definedName><definedName name="Secret" hidden="1" localSheetId="0"><![CDATA[IF(1<2,"&amp;",0)]]></definedName><definedName name="Empty" hidden="true"/><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName><definedName name="Unicode">&quot;&#26679;&amp;&quot;</definedName></definedNames><calcPr calcMode="manual" iterate="1" iterateCount="25" fullPrecision="0"/><externalReferences><externalReference r:id="rIdExternal"/></externalReferences><foreign:definedName name="Imposter">99</foreign:definedName>"#, prefix).unwrap();
            assert_eq!(context.names.len(), 6);
            assert_eq!(context.names[0].sheet_id, None);
            assert_eq!(context.names[1].sheet_id.as_deref(), Some("sheet-2"));
            assert_eq!(context.names[2].sheet_id.as_deref(), Some("sheet-1"));
            assert!(context.names[2].hidden);
            assert_eq!(context.names[2].formula, "IF(1<2,\"&amp;\",0)");
            assert_eq!(context.names[3].formula, "");
            assert_eq!(context.names[5].formula, "\"样&\"");
            assert_eq!(context.calculation_properties["calcMode"], "manual");
            assert_eq!(context.calculation_properties["iterateCount"], "25");
            assert_eq!(context.external_relationship_ids, vec!["rIdExternal"]);
        }
    }

    #[test]
    fn rejects_invalid_scope_and_external_relationship_namespace() {
        for body in [
            r#"<definedNames><definedName name="Secret" hidden="1" localSheetId="9">1</definedName></definedNames>"#,
            r#"<definedNames><definedName name="Secret" hidden="1" localSheetId="bad">1</definedName></definedNames>"#,
            r#"<externalReferences><externalReference foreign:id="fake"/></externalReferences>"#,
            r#"<definedNames><definedName name="Secret" hidden="1"><foreign:text>1</foreign:text></definedName></definedNames>"#,
        ] {
            assert!(read(body, false).is_err(), "{body}");
        }
    }

    #[test]
    fn closed_context_reads_fail() {
        let sessions = WorkbookSessions::new();
        assert!(
            sessions
                .read_calculation_context("closed", &AtomicBool::new(false))
                .is_err()
        );
    }
}
