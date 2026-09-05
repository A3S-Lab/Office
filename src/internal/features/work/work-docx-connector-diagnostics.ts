import { attribute, descendants } from './work-ooxml-package';

/**
 * WPS Writer writes connectors as legacy VML shapes. The importer owns the
 * isolated straight subset; this inspection remains package-level evidence so
 * diagnostics can explain which native records may normalize on edit.
 */
export interface DocxConnectorInspection {
  detected: number;
  connectorPictContainers: number;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const VML_NAMESPACE = 'urn:schemas-microsoft-com:vml';
const OFFICE_NAMESPACE = 'urn:schemas-microsoft-com:office:office';

/**
 * Counts only VML connector records, not ordinary VML pictures or text
 * boxes. WPS emits shape types 32 (straight), 33 (elbow), and 37 (curved)
 * for connectors created by `Shapes.AddConnector`.
 */
export function inspectDocxConnectors(
  document: Document,
): DocxConnectorInspection {
  let detected = 0;
  let connectorPictContainers = 0;
  for (const pict of descendants(document, 'pict')) {
    if (!isWordElement(pict)) continue;
    const vmlShapes = [
      ...descendants(pict, 'shape'),
      ...descendants(pict, 'line'),
      ...descendants(pict, 'polyline'),
    ].filter((element) => element.namespaceURI === VML_NAMESPACE);
    const connectors = vmlShapes.filter(isVmlConnector);
    detected += connectors.length;
    if (connectors.length > 0 && connectors.length === vmlShapes.length) {
      connectorPictContainers += 1;
    }
  }
  return { detected, connectorPictContainers };
}

function isVmlConnector(element: Element): boolean {
  if (element.namespaceURI !== VML_NAMESPACE) return false;
  if (element.localName === 'line' || element.localName === 'polyline') {
    return true;
  }
  const shapeType =
    element.getAttributeNS(OFFICE_NAMESPACE, 'spt') ??
    attribute(element, 'o:spt');
  const type = attribute(element, 'type');
  return (
    ['32', '33', '37'].includes(shapeType?.trim() ?? '') ||
    ['#_x0000_t32', '#_x0000_t33', '#_x0000_t37'].includes(
      type?.trim().toLowerCase() ?? '',
    )
  );
}

function isWordElement(element: Element): boolean {
  return (
    element.namespaceURI === WORD_NAMESPACE ||
    element.namespaceURI === STRICT_WORD_NAMESPACE
  );
}
