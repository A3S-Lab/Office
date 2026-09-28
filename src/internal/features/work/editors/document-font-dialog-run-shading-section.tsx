import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type { DocumentRunShadingPattern } from '../work-document-run-shading';
import type {
  DocumentFontDialogRunShadingDraft,
  DocumentFontDialogRunShadingMode,
  DocumentFontDialogRunShadingSource,
} from './document-font-dialog-run-shading-model';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

function shadingModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentFontDialogRunShadingMode;
  label: string;
  disabled?: boolean;
}> {
  return [
    {
      value: 'mixed',
      label: officeMessage(messages, 'document.font.mode.mixed'),
      disabled: true,
    },
    {
      value: 'inherit',
      label: officeMessage(messages, 'document.font.mode.inherit'),
    },
    {
      value: 'none',
      label: officeMessage(messages, 'document.font.mode.noneExplicit'),
    },
    {
      value: 'value',
      label: officeMessage(messages, 'document.font.shading.modeValue'),
    },
  ];
}

function shadingPatterns(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: Exclude<DocumentRunShadingPattern, 'nil'>;
  label: string;
  group?: string;
}> {
  const basic = officeMessage(messages, 'document.font.shading.group.basic');
  const stripes = officeMessage(
    messages,
    'document.font.shading.group.stripes',
  );
  const cross = officeMessage(messages, 'document.font.shading.group.cross');
  const thin = officeMessage(messages, 'document.font.shading.group.thin');
  const density = officeMessage(
    messages,
    'document.font.shading.group.density',
  );
  return [
    {
      value: 'clear',
      label: officeMessage(messages, 'document.font.shading.pattern.clear'),
      group: basic,
    },
    {
      value: 'solid',
      label: officeMessage(messages, 'document.font.shading.pattern.solid'),
      group: basic,
    },
    {
      value: 'horzStripe',
      label: officeMessage(messages, 'document.font.shading.pattern.horzStripe'),
      group: stripes,
    },
    {
      value: 'vertStripe',
      label: officeMessage(messages, 'document.font.shading.pattern.vertStripe'),
      group: stripes,
    },
    {
      value: 'reverseDiagStripe',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.reverseDiagStripe',
      ),
      group: stripes,
    },
    {
      value: 'diagStripe',
      label: officeMessage(messages, 'document.font.shading.pattern.diagStripe'),
      group: stripes,
    },
    {
      value: 'horzCross',
      label: officeMessage(messages, 'document.font.shading.pattern.horzCross'),
      group: cross,
    },
    {
      value: 'diagCross',
      label: officeMessage(messages, 'document.font.shading.pattern.diagCross'),
      group: cross,
    },
    {
      value: 'thinHorzStripe',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.thinHorzStripe',
      ),
      group: thin,
    },
    {
      value: 'thinVertStripe',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.thinVertStripe',
      ),
      group: thin,
    },
    {
      value: 'thinReverseDiagStripe',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.thinReverseDiagStripe',
      ),
      group: thin,
    },
    {
      value: 'thinDiagStripe',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.thinDiagStripe',
      ),
      group: thin,
    },
    {
      value: 'thinHorzCross',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.thinHorzCross',
      ),
      group: thin,
    },
    {
      value: 'thinDiagCross',
      label: officeMessage(
        messages,
        'document.font.shading.pattern.thinDiagCross',
      ),
      group: thin,
    },
    ...[
      5, 10, 12, 15, 20, 25, 30, 35, 37, 40, 45, 50, 55, 60, 62, 65, 70, 75, 80,
      85, 87, 90, 95,
    ].map((percentage) => ({
      value: `pct${percentage}` as Exclude<DocumentRunShadingPattern, 'nil'>,
      label: `${percentage}%`,
      group: density,
    })),
  ];
}

export function DocumentFontDialogRunShadingSection({
  source,
  draft,
  touched,
  onDraftChange,
  onTouched,
}: {
  source: DocumentFontDialogRunShadingSource;
  draft: DocumentFontDialogRunShadingDraft;
  touched: boolean;
  onDraftChange: (patch: Partial<DocumentFontDialogRunShadingDraft>) => void;
  onTouched: () => void;
}) {
  const messages = useOfficeMessages();
  const enabled = draft.runShadingMode === 'value';
  const update = (patch: Partial<DocumentFontDialogRunShadingDraft>) => {
    onDraftChange(patch);
    onTouched();
  };
  return (
    <fieldset
      className="work-document-font-dialog-run-shading"
      aria-label={officeMessage(messages, 'document.font.shading.sectionAria')}
    >
      <legend>
        {officeMessage(messages, 'document.font.shading.legend')}
      </legend>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.shading.apply')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'document.font.shading.applyAria')}
          value={draft.runShadingMode}
          options={shadingModes(messages)}
          onValueChange={(runShadingMode) => update({ runShadingMode })}
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.shading.pattern')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(
            messages,
            'document.font.shading.patternAria',
          )}
          value={draft.runShadingPattern}
          options={shadingPatterns(messages)}
          disabled={!enabled}
          onValueChange={(runShadingPattern) =>
            update({ runShadingMode: 'value', runShadingPattern })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>
          {officeMessage(messages, 'document.font.shading.foreground')}
        </span>
        <OfficeColorPicker
          ariaLabel={officeMessage(
            messages,
            'document.font.shading.foregroundAria',
          )}
          value={
            draft.runShadingColor === 'auto' ? '#000000' : draft.runShadingColor
          }
          disabled={!enabled}
          resetAction={{
            kind: 'automatic',
            label: officeMessage(
              messages,
              'document.font.shading.foregroundAuto',
            ),
            onSelect: () =>
              update({ runShadingMode: 'value', runShadingColor: 'auto' }),
          }}
          onValueChange={(runShadingColor) =>
            update({
              runShadingMode: 'value',
              runShadingColor: runShadingColor as `#${string}`,
            })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>
          {officeMessage(messages, 'document.font.shading.background')}
        </span>
        <OfficeColorPicker
          ariaLabel={officeMessage(
            messages,
            'document.font.shading.backgroundAria',
          )}
          value={
            draft.runShadingFill === 'auto' ? '#ffffff' : draft.runShadingFill
          }
          disabled={!enabled}
          resetAction={{
            kind: 'automatic',
            label: officeMessage(
              messages,
              'document.font.shading.backgroundAuto',
            ),
            onSelect: () =>
              update({ runShadingMode: 'value', runShadingFill: 'auto' }),
          }}
          onValueChange={(runShadingFill) =>
            update({
              runShadingMode: 'value',
              runShadingFill: runShadingFill as `#${string}`,
            })
          }
        />
      </div>
      {source.mixed && !touched ? (
        <p className="work-document-font-dialog-mixed" role="status">
          {officeMessage(messages, 'document.font.shading.mixed')}
        </p>
      ) : null}
    </fieldset>
  );
}
