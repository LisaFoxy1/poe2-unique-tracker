import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { fetch } from "@tauri-apps/plugin-http";
import { readFile } from "@tauri-apps/plugin-fs";
import Database from "@tauri-apps/plugin-sql";
import { emit, listen } from "@tauri-apps/api/event";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import {
  cursorPosition,
  LogicalPosition,
  monitorFromPoint,
  primaryMonitor,
} from "@tauri-apps/api/window";
import { invoke } from "@tauri-apps/api/core";
import { readText, writeText } from "@tauri-apps/plugin-clipboard-manager";
import { register, unregister } from "@tauri-apps/plugin-global-shortcut";
import { read, utils } from "xlsx";
import {
  ensureEditionAvailabilitySchema,
  isCanonicalCatalogueReady,
  reconcileImportedCollection,
  shouldCheckCatalogueNow,
  syncCatalogueFromPoeWiki,
  type EditionAvailability,
  type ImportReconciliationSummary,
} from "./catalogueSync";
import {
  identifyUniqueFromClipboard,
  type ItemIdentificationResult,
} from "./poeItemParser";
import ImportReviewModal from "./ImportReviewModal";
import AppUpdater from "./AppUpdater";
import "./App.css";

type TrackingFlag =
  | "owned"
  | "wearing"
  | "foil"
  | "foulborn"
  | "vestigial";

type ExtraVariant = "foil" | "foulborn" | "vestigial";

type EditionAvailabilityMap = Record<
  ExtraVariant,
  EditionAvailability
>;

type EditionSourceMap = Record<
  ExtraVariant,
  string
>;

type UniqueEntry = {
  id: string;
  familyId: string;
  name: string;
  baseType: string | null;
  itemType: string;
  variantLabel: string | null;
  importStatus: string;
  flags: TrackingFlag[];
  editionAvailability: EditionAvailabilityMap;
  editionSources: EditionSourceMap;
  reviewed: boolean;
  isLegacyOnly: boolean;
  catalogueSource: string;
};

type CatalogueUpdateSummary = {
  revision: string;
  label: string;
  newFamilies: number;
  newVariants: number;
  dropDisabled: number;
  updatedEntries: number;
};

type TrackerDiscoveryItem = {
  id: string;
  name: string;
  itemType: string;
  variantLabel: string | null;
  releaseVersion: string | null;
};

type PendingLeagueDetection = {
  profileId: string;
  leagueName: string;
};

type TrackerDiscoveryNotice =
  | {
      id: string;
      kind: "new-league";
      leagueName: string;
      profileId: string;
      items: TrackerDiscoveryItem[];
      createdAt: number;
    }
  | {
      id: string;
      kind: "new-unique";
      items: TrackerDiscoveryItem[];
      createdAt: number;
    };

type SortMode = "alphabetical" | "type";

type StatusFilter =
  | "all"
  | "unreviewed"
  | "missing"
  | TrackingFlag;

type StatusKey = "missing" | TrackingFlag;

type StatusColors = Record<StatusKey, string>;

type CollectionRules = Record<TrackingFlag, boolean>;

type ExtraTracking = Record<ExtraVariant, boolean>;

type CollectionProfile = {
  id: string;
  name: string;
  kind: "standard" | "challenge";
  isArchived: boolean;
};

type PoeTradeLeagueEntry = {
  id: string;
  text: string;
  realm?: string | null;
};

type PoeTradeLeagueResponse = {
  result: PoeTradeLeagueEntry[];
};

type WikiPageTitleResponse = {
  query?: {
    pages?: Array<{
      pageid?: number;
      title?: string;
    }>;
  };
};

type LeagueRolloverChange = {
  uniqueId: string;
  name: string;
  addedFlags: TrackingFlag[];
};

type LeagueRolloverPreview = {
  profileId: string;
  oldLeagueName: string;
  newLeagueName: string;
  changedUniques: number;
  flagCounts: Record<TrackingFlag, number>;
  changes: LeagueRolloverChange[];
};

type LeagueRolloverMode =
  | "dev-pending"
  | "dev-complete"
  | "pending"
  | "complete";

type SqliteTransactionStatement = {
  sql: string;
  params?: string[];
};

type ImportMode =
  | "status-list"
  | "missing-only"
  | "color-coded-list";

type ColorCodedStatus =
  | "missing"
  | "league-owned"
  | "standard-owned"
  | "unreviewed";

type ParsedImportRow = {
  id: string;
  name: string;
  itemType: string;
  rawStatus: string;
  sourceSheetName: string;
  sourceRowIndex: number;
  colorCodedStatus: ColorCodedStatus;
};

type ProtectedLeagueUnique = {
  id: string;
  name: string;
  itemType: string;
  variantLabel: string | null;
  releaseVersion: string | null;
};

type PendingImport = {
  fileName: string;
  rows: ParsedImportRow[];
  suggestedMode: ImportMode;
  latestReleaseLine: string | null;
  protectedLeagueUniques: ProtectedLeagueUnique[];
};

type MissingOnlyInferenceSummary = {
  kind: "missing-only";
  inferredOwned: number;
  explicitMissing: number;
  protectedNewLeague: ProtectedLeagueUnique[];
  protectedAmbiguous: number;
  latestReleaseLine: string | null;
  destinationProfileId: string;
  destinationProfileName: string;
};

type ColorCodedImportSummary = {
  kind: "color-coded-list";
  standardOwned: number;
  leagueOwned: number;
  missing: number;
  unknownCandidatesMarkedUnreviewed: number;
  unresolvedPlaceholders: number;
  unresolvedNamedRows: number;
  leagueProfileId: string;
  leagueProfileName: string;
};

type ImportApplicationSummary =
  | MissingOnlyInferenceSummary
  | ColorCodedImportSummary;

type SheetJsColor = {
  rgb?: string;
};

type SheetJsStyledCell = {
  s?: {
    /*
     * SheetJS can expose parsed styles in either its flat
     * style shape or the nested font/fill shape depending
     * on the workbook/style path. Support both.
     */
    color?: SheetJsColor;
    fgColor?: SheetJsColor;
    font?: {
      color?: SheetJsColor;
    };
    fill?: {
      fgColor?: SheetJsColor;
    };
  };
};

const COLOR_CODED_IMPORT_STATUS = {
  missing: "Color-coded Missing",
  leagueOwned: "Color-coded League Owned",
  standardOwned: "Color-coded Standard Owned",
} as const;

const STANDARD_PROFILE_ID = "standard";
const CURRENT_LEAGUE_PROFILE_ID = "current-league";

const PENDING_DISCOVERY_NOTICES_META_KEY =
  "pending_tracker_discovery_notices_v1";

const PENDING_NEW_LEAGUES_META_KEY =
  "pending_new_league_detections_v1";

function makeLeagueProfileId(
  leagueKey: string,
) {
  return `league:${encodeURIComponent(
    leagueKey,
  )}`;
}

const IMPORT_SHEET_TYPES: Record<string, string> = {
  flask: "Flask",
  flasks: "Flask",
  amulet: "Amulet",
  amulets: "Amulet",
  ring: "Ring",
  rings: "Ring",
  wand: "Wand",
  wands: "Wand",
  mace: "Mace",
  maces: "Mace",
  "one hand mace": "Mace",
  "one hand maces": "Mace",
  "two hand mace": "Mace",
  "two hand maces": "Mace",
  bow: "Bow",
  bows: "Bow",
  staff: "Staff",
  staves: "Staff",
  quiver: "Quiver",
  quivers: "Quiver",
  belt: "Belt",
  belts: "Belt",
  glove: "Gloves",
  gloves: "Gloves",
  boot: "Boots",
  boots: "Boots",
  "body armour": "Body Armour",
  "body armours": "Body Armour",
  helmet: "Helmet",
  helmets: "Helmet",
  shield: "Shield",
  shields: "Shield",
  buckler: "Shield",
  bucklers: "Shield",
  targe: "Shield",
  targes: "Shield",
  jewel: "Jewel",
  jewels: "Jewel",
  charm: "Charm",
  charms: "Charm",
  crossbow: "Crossbow",
  crossbows: "Crossbow",
  focus: "Focus",
  foci: "Focus",
  sceptre: "Sceptre",
  sceptres: "Sceptre",
  spear: "Spear",
  spears: "Spear",
  quarterstaff: "Quarterstaff",
  quarterstaves: "Quarterstaff",
  talisman: "Talisman",
  talismans: "Talisman",
  flail: "Flail",
  flails: "Flail",
  sword: "Sword",
  swords: "Sword",
  axe: "Axe",
  axes: "Axe",
  dagger: "Dagger",
  daggers: "Dagger",
  waystone: "Waystone",
  waystones: "Waystone",
  tablet: "Tablet",
  tablets: "Tablet",
  relic: "Relic",
  relics: "Relic",
};

const POE2_UNIQUE_TAB_TYPE_ORDER = [
  "Flask",
  "Amulet",
  "Ring",
  "Wand",
  "Mace",
  "Bow",
  "Staff",
  "Quiver",
  "Belt",
  "Gloves",
  "Boots",
  "Body Armour",
  "Helmet",
  "Shield",
  "Jewel",
  "Charm",
  "Crossbow",
  "Focus",
  "Sceptre",
  "Spear",
  "Quarterstaff",
  "Talisman",
] as const;

const POE2_UNIQUE_TAB_TYPE_RANK =
  new Map<string, number>(
    POE2_UNIQUE_TAB_TYPE_ORDER.map(
      (itemType, index) => [
        itemType,
        index,
      ],
    ),
  );

function comparePoe2UniqueTabTypes(
  left: string,
  right: string,
) {
  if (left === right) {
    return 0;
  }

  /*
   * Sword uniques arrive with 0.5.5, but we do not yet
   * know where GGG places Swords in the unique stash tab.
   * Keep them at the very bottom until the live tab order
   * can be confirmed.
   */
  if (left === "Sword") {
    return 1;
  }

  if (right === "Sword") {
    return -1;
  }

  const leftRank =
    POE2_UNIQUE_TAB_TYPE_RANK.get(
      left,
    );
  const rightRank =
    POE2_UNIQUE_TAB_TYPE_RANK.get(
      right,
    );

  if (
    leftRank !== undefined &&
    rightRank !== undefined
  ) {
    return leftRank - rightRank;
  }

  if (leftRank !== undefined) {
    return -1;
  }

  if (rightRank !== undefined) {
    return 1;
  }

  /*
   * Future or otherwise unlisted item types remain stable
   * and usable: they come after the known stash-tab types,
   * alphabetically, but still before Sword.
   */
  return left.localeCompare(right);
}

const IMPORT_NAME_ALIASES = new Map([
  ["advanced fortress", "advancing fortress"],
]);

const KNOWN_IMPORT_STATUSES = new Set([
  "owned",
  "wearing",
  "owned foil",
  "owned rainbow",
  "missing",
  "unreviewed",
]);

function normalizeImportItemType(sheetName: string) {
  return (
    IMPORT_SHEET_TYPES[
      sheetName.trim().toLowerCase()
    ] ?? null
  );
}

function normalizeImportName(value: string) {
  return value
    .normalize("NFKC")
    .replace(/[\u2019\u2018]/g, "'")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function importNameWithAlias(value: string) {
  const normalized = normalizeImportName(value);
  return IMPORT_NAME_ALIASES.get(normalized) ?? normalized;
}

function normalizeSheetJsRgb(
  value: string | undefined,
) {
  if (!value) {
    return null;
  }

  const normalized = value
    .replace(/[^0-9a-f]/gi, "")
    .toUpperCase();

  return normalized.length >= 6
    ? normalized.slice(-6)
    : null;
}

function getColorCodedStatus(
  cell: SheetJsStyledCell | undefined,
  name: string,
): ColorCodedStatus {
  const compactName =
    name.replace(/\s+/g, "");

  if (/^\?+$/.test(compactName)) {
    return "unreviewed";
  }

  const fontColor =
    normalizeSheetJsRgb(
      cell?.s?.font?.color?.rgb ??
        cell?.s?.color?.rgb,
    );

  const fillColor =
    normalizeSheetJsRgb(
      cell?.s?.fill?.fgColor?.rgb ??
        cell?.s?.fgColor?.rgb,
    );

  if (
    fontColor === "006100" &&
    fillColor === "C6EFCE"
  ) {
    return "standard-owned";
  }

  if (fontColor === "00B050") {
    return "league-owned";
  }

  return "missing";
}


function importNameDistance(leftValue: string, rightValue: string) {
  const left = importNameWithAlias(leftValue);
  const right = importNameWithAlias(rightValue);

  if (left === right) {
    return 0;
  }

  if (!left) {
    return right.length;
  }

  if (!right) {
    return left.length;
  }

  let previous = Array.from(
    { length: right.length + 1 },
    (_, index) => index,
  );

  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];

    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      current[rightIndex] = Math.min(
        current[rightIndex - 1] + 1,
        previous[rightIndex] + 1,
        previous[rightIndex - 1] +
          (left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1),
      );
    }

    previous = current;
  }

  return previous[right.length];
}

function isLikelySameImportName(left: string, right: string) {
  const normalizedLeft = importNameWithAlias(left);
  const normalizedRight = importNameWithAlias(right);

  if (!normalizedLeft || !normalizedRight) {
    return false;
  }

  if (normalizedLeft === normalizedRight) {
    return true;
  }

  const longest = Math.max(
    normalizedLeft.length,
    normalizedRight.length,
  );
  const distance = importNameDistance(
    normalizedLeft,
    normalizedRight,
  );

  if (longest < 8) {
    return distance <= 1;
  }

  if (longest < 20) {
    return distance <= 2;
  }

  return distance <= 3 && distance / longest <= 0.12;
}

function releaseLine(value: string | null | undefined) {
  const match = /^(\d+)\.(\d+)/.exec(value?.trim() ?? "");

  if (!match) {
    return null;
  }

  return `${Number(match[1])}.${Number(match[2])}`;
}

function latestReleaseLine(values: Array<string | null>) {
  let best: { major: number; minor: number; label: string } | null = null;

  for (const value of values) {
    const match = /^(\d+)\.(\d+)/.exec(value?.trim() ?? "");

    if (!match) {
      continue;
    }

    const major = Number(match[1]);
    const minor = Number(match[2]);

    if (
      !best ||
      major > best.major ||
      (major === best.major && minor > best.minor)
    ) {
      best = {
        major,
        minor,
        label: `${major}.${minor}`,
      };
    }
  }

  return best?.label ?? null;
}

async function readJsonMeta<T>(
  db: Database,
  key: string,
  fallback: T,
): Promise<T> {
  const rows = await db.select<
    { value: string }[]
  >(
    `
      SELECT value
      FROM app_meta
      WHERE key = ?
    `,
    [key],
  );

  const raw = rows[0]?.value;

  if (!raw) {
    return fallback;
  }

  try {
    return JSON.parse(raw) as T;
  } catch (error) {
    console.warn(
      `Could not parse saved tracker metadata for ${key}:`,
      error,
    );

    return fallback;
  }
}

async function writeJsonMeta(
  db: Database,
  key: string,
  value: unknown,
) {
  await db.execute(
    `
      INSERT OR REPLACE INTO app_meta (
        key,
        value
      )
      VALUES (?, ?)
    `,
    [
      key,
      JSON.stringify(value),
    ],
  );
}

async function readPendingLeagueDetections(
  db: Database,
): Promise<PendingLeagueDetection[]> {
  const saved =
    await readJsonMeta<unknown>(
      db,
      PENDING_NEW_LEAGUES_META_KEY,
      [],
    );

  if (!Array.isArray(saved)) {
    return [];
  }

  return saved.flatMap((entry) => {
    if (
      !entry ||
      typeof entry !== "object"
    ) {
      return [];
    }

    const value = entry as {
      profileId?: unknown;
      leagueName?: unknown;
    };

    if (
      typeof value.profileId !== "string" ||
      typeof value.leagueName !== "string"
    ) {
      return [];
    }

    return [
      {
        profileId: value.profileId,
        leagueName: value.leagueName,
      },
    ];
  });
}

async function appendPendingLeagueDetections(
  db: Database,
  additions: PendingLeagueDetection[],
) {
  if (additions.length === 0) {
    return;
  }

  const existing =
    await readPendingLeagueDetections(
      db,
    );

  const byProfileId =
    new Map<
      string,
      PendingLeagueDetection
    >(
      existing.map((entry) => [
        entry.profileId,
        entry,
      ]),
    );

  for (const addition of additions) {
    byProfileId.set(
      addition.profileId,
      addition,
    );
  }

  await writeJsonMeta(
    db,
    PENDING_NEW_LEAGUES_META_KEY,
    Array.from(
      byProfileId.values(),
    ),
  );
}

async function clearPendingLeagueDetections(
  db: Database,
) {
  await writeJsonMeta(
    db,
    PENDING_NEW_LEAGUES_META_KEY,
    [],
  );
}

async function readDiscoveryNotices(
  db: Database,
): Promise<TrackerDiscoveryNotice[]> {
  const saved =
    await readJsonMeta<unknown>(
      db,
      PENDING_DISCOVERY_NOTICES_META_KEY,
      [],
    );

  if (!Array.isArray(saved)) {
    return [];
  }

  const notices:
    TrackerDiscoveryNotice[] = [];

  for (const entry of saved) {
    if (
      !entry ||
      typeof entry !== "object"
    ) {
      continue;
    }

    const value = entry as {
      id?: unknown;
      kind?: unknown;
      leagueName?: unknown;
      profileId?: unknown;
      items?: unknown;
      createdAt?: unknown;
    };

    if (
      typeof value.id !== "string" ||
      (value.kind !== "new-league" &&
        value.kind !== "new-unique") ||
      !Array.isArray(value.items)
    ) {
      continue;
    }

    const items:
      TrackerDiscoveryItem[] = [];

    for (const item of value.items) {
      if (
        !item ||
        typeof item !== "object"
      ) {
        continue;
      }

      const candidate = item as {
        id?: unknown;
        name?: unknown;
        itemType?: unknown;
        variantLabel?: unknown;
        releaseVersion?: unknown;
      };

      if (
        typeof candidate.id !== "string" ||
        typeof candidate.name !== "string" ||
        typeof candidate.itemType !== "string"
      ) {
        continue;
      }

      items.push({
        id: candidate.id,
        name: candidate.name,
        itemType:
          candidate.itemType,
        variantLabel:
          typeof candidate.variantLabel ===
          "string"
            ? candidate.variantLabel
            : null,
        releaseVersion:
          typeof candidate.releaseVersion ===
          "string"
            ? candidate.releaseVersion
            : null,
      });
    }

    const createdAt =
      typeof value.createdAt === "number" &&
      Number.isFinite(value.createdAt)
        ? value.createdAt
        : Date.now();

    if (value.kind === "new-league") {
      if (
        typeof value.leagueName !== "string" ||
        typeof value.profileId !== "string"
      ) {
        continue;
      }

      notices.push({
        id: value.id,
        kind: "new-league",
        leagueName:
          value.leagueName,
        profileId:
          value.profileId,
        items,
        createdAt,
      });

      continue;
    }

    notices.push({
      id: value.id,
      kind: "new-unique",
      items,
      createdAt,
    });
  }

  return notices;
}

async function appendDiscoveryNotices(
  db: Database,
  additions: TrackerDiscoveryNotice[],
) {
  const existing =
    await readDiscoveryNotices(
      db,
    );

  const byId =
    new Map<
      string,
      TrackerDiscoveryNotice
    >(
      existing.map((notice) => [
        notice.id,
        notice,
      ]),
    );

  for (const addition of additions) {
    byId.set(
      addition.id,
      addition,
    );
  }

  const merged =
    Array.from(
      byId.values(),
    ).sort(
      (left, right) =>
        left.createdAt -
        right.createdAt,
    );

  await writeJsonMeta(
    db,
    PENDING_DISCOVERY_NOTICES_META_KEY,
    merged,
  );

  return merged;
}

async function saveDiscoveryNotices(
  db: Database,
  notices: TrackerDiscoveryNotice[],
) {
  await writeJsonMeta(
    db,
    PENDING_DISCOVERY_NOTICES_META_KEY,
    notices,
  );
}

const DEFAULT_STATUS_COLORS: StatusColors = {
  missing: "#d76868",
  owned: "#79ba68",
  wearing: "#71a5dc",
  foil: "rainbow",
  foulborn: "#c99c72",
  vestigial: "#b79bd7",
};

const STATUS_LABELS: Record<StatusKey, string> = {
  missing: "Missing",
  owned: "Owned",
  wearing: "Wearing",
  foil: "Foil",
  foulborn: "Foulborn",
  vestigial: "Vestigial",
};

const STATUS_KEYS: StatusKey[] = [
  "missing",
  "owned",
  "wearing",
  "foil",
];

const DEFAULT_COLLECTION_RULES: CollectionRules = {
  owned: true,
  wearing: true,
  foil: true,
  foulborn: true,
  vestigial: true,
};

const DEFAULT_EXTRA_TRACKING: ExtraTracking = {
  foil: true,
  foulborn: false,
  vestigial: false,
};

const EXTRA_VARIANTS: ExtraVariant[] = [
  "foil",
];

const DEFAULT_EDITION_AVAILABILITY: EditionAvailabilityMap = {
  foil: "unknown",
  foulborn: "unavailable",
  vestigial: "unavailable",
};

const DEFAULT_EDITION_SOURCES: EditionSourceMap = {
  foil: "default",
  foulborn: "default",
  vestigial: "default",
};

const CONFIDENT_EDITION_SOURCES = new Set([
  "manual",
  "poewiki-foulborn-category",
  "poewiki-foil-confirmed",
  "built-in-vestigial-class-rule",
  "poe2-reliquary-general",
]);

function isExtraVariant(
  value: string,
): value is ExtraVariant {
  return EXTRA_VARIANTS.includes(
    value as ExtraVariant,
  );
}

function isEditionAvailability(
  value: string,
): value is EditionAvailability {
  return (
    value === "available" ||
    value === "unavailable" ||
    value === "unknown"
  );
}

function shouldShowEdition(
  edition: ExtraVariant,
  availability: EditionAvailabilityMap,
  flags: TrackingFlag[],
  extraTracking: ExtraTracking,
  detectedEdition?:
    | "normal"
    | "foulborn"
    | "vestigial",
) {
  if (!extraTracking[edition]) {
    return false;
  }

  /*
   * Existing player data always wins. Never hide a
   * status somebody already has saved.
   *
   * Likewise, if the item parser literally detects a
   * Foulborn/Vestigial item, that is direct evidence
   * that the edition exists.
   */
  if (
    flags.includes(edition) ||
    detectedEdition === edition
  ) {
    return true;
  }

  /*
   * Unknown stays visible until we have enough evidence
   * to call an edition impossible.
   */
  return availability[edition] !== "unavailable";
}

function isEditionUncertain(
  edition: ExtraVariant,
  availability: EditionAvailabilityMap,
  sources: EditionSourceMap,
  flags: TrackingFlag[],
  detectedEdition?:
    | "normal"
    | "foulborn"
    | "vestigial",
) {
  /*
   * Direct evidence always wins.
   *
   * If the parser literally sees the edition in-game,
   * or the player says they own one, PoE 2 Unique Tracker
   * treats that edition as confirmed to exist.
   */
  if (
    detectedEdition === edition ||
    flags.includes(edition)
  ) {
    return false;
  }

  if (
    availability[edition] === "unknown"
  ) {
    return true;
  }

  if (
    availability[edition] === "available"
  ) {
    return !CONFIDENT_EDITION_SOURCES.has(
      sources[edition],
    );
  }

  return false;
}

function isTrackingFlag(value: string): value is TrackingFlag {
  return [
    "owned",
    "wearing",
    "foil",
    "foulborn",
    "vestigial",
  ].includes(value);
}

function statusToFlags(status: string): TrackingFlag[] {
  switch (status) {
    case "Owned":
      return ["owned"];
    case "Wearing":
      return ["wearing"];
    case "Owned Foil":
    case "Owned Rainbow":
      return ["owned", "foil"];
    default:
      return [];
  }
}

function getStatusStyle(color: string): CSSProperties | undefined {
  if (color === "rainbow") {
    return undefined;
  }

  return {
    "--status-color": color,
  } as CSSProperties;
}

function makeFamilyId(name: string) {
  const slug = name
    .normalize("NFKD")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return `family:${slug || "unknown"}`;
}

function getPoewikiPageId(
  uniqueId: string,
) {
  const match =
    /^poewiki:(\d+):/.exec(
      uniqueId,
    );

  return match?.[1] ?? null;
}

function extractWikiVariantLabel(
  itemName: string,
  pageTitle: string,
) {
  const escapedName =
    itemName.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&",
    );

  const match =
    pageTitle.match(
      new RegExp(
        `^${escapedName} \\((.+)\\)$`,
        "i",
      ),
    );

  return (
    match?.[1]?.trim() ||
    null
  );
}

const SYNTHESIS_VARIANT_FAMILIES =
  new Set([
    "Circle of Anguish",
    "Circle of Fear",
    "Circle of Guilt",
    "Circle of Nostalgia",
    "Circle of Regret",
    "Garb of the Ephemeral",
    "Mask of the Tribunal",
    "Nebulis",
    "Offering to the Serpent",
    "Perepiteia",
  ]);

function getCuratedVariantLabel(
  itemName: string,
  wikiLabel: string | null,
) {
  // PoE 1 had several hand-curated name-specific rules.
  // PoE 2 uses the Wiki label directly.
  void itemName;
  void SYNTHESIS_VARIANT_FAMILIES;
  return wikiLabel;
}

function isCuratedLegacyVariant(
  itemName: string,
  variantLabel: string | null,
) {
  // Legacy state for PoE 2 comes from catalogue metadata rather
  // than PoE 1-specific hand-written item rules.
  void itemName;
  void variantLabel;
  return false;
}

async function fetchWikiVariantLabels(
  rows: Array<{
    id: string;
    name: string;
  }>,
) {
  /*
   * Only ask the Wiki about families that actually have
   * multiple catalogue rows. Single-version uniques
   * don't need variant labels.
   */
  const familyCounts =
    new Map<string, number>();

  for (const row of rows) {
    const familyId =
      makeFamilyId(row.name);

    familyCounts.set(
      familyId,
      (familyCounts.get(
        familyId,
      ) ?? 0) + 1,
    );
  }

  const candidates =
    rows.filter(
      (row) =>
        (
          familyCounts.get(
            makeFamilyId(
              row.name,
            ),
          ) ?? 0
        ) > 1,
    );

  const pageIds =
    Array.from(
      new Set(
        candidates
          .map((row) =>
            getPoewikiPageId(
              row.id,
            ),
          )
          .filter(
            (
              pageId,
            ): pageId is string =>
              pageId !== null,
          ),
      ),
    );

  const titleByPageId =
    new Map<
      string,
      string
    >();

  /*
   * MediaWiki accepts page IDs in batches. Keeping this
   * at 50 avoids relying on elevated API limits.
   */
  const BATCH_SIZE = 50;

  for (
    let index = 0;
    index < pageIds.length;
    index += BATCH_SIZE
  ) {
    const batch =
      pageIds.slice(
        index,
        index + BATCH_SIZE,
      );

    const params =
      new URLSearchParams({
        action: "query",
        format: "json",
        formatversion: "2",
        origin: "*",
        pageids:
          batch.join("|"),
      });

    try {
      const response =
        await fetch(
          `https://www.poe2wiki.net/w/api.php?${params.toString()}`,
          {
            method: "GET",
            headers: {
              Accept:
                "application/json",
            },
          },
        );

      if (!response.ok) {
        console.warn(
          `Variant title lookup failed with HTTP ${response.status}.`,
        );

        continue;
      }

      const payload =
        (await response.json()) as
          WikiPageTitleResponse;

      for (
        const page of
        payload.query?.pages ??
        []
      ) {
        if (
          page.pageid == null ||
          !page.title
        ) {
          continue;
        }

        titleByPageId.set(
          String(page.pageid),
          page.title,
        );
      }
    } catch (error) {
      /*
       * Variant labels are optional metadata.
       * Never stop the local catalogue from loading
       * just because the Wiki is unavailable.
       */
      console.warn(
        "Could not fetch PoE 2 Wiki variant titles:",
        error,
      );
    }
  }

  const labels =
    new Map<
      string,
      string
    >();

  for (const row of candidates) {
    const pageId =
      getPoewikiPageId(
        row.id,
      );

    if (!pageId) {
      continue;
    }

    const pageTitle =
      titleByPageId.get(
        pageId,
      );

    if (!pageTitle) {
      continue;
    }

    const label =
      extractWikiVariantLabel(
        row.name,
        pageTitle,
      );

    if (label) {
      labels.set(
        row.id,
        label,
      );
    }
  }

  return labels;
}

function isExcludedLeagueKey(
  value: string,
) {
  const normalized =
    value
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();

  return (
    normalized === "standard" ||
    normalized === "hardcore" ||
    normalized === "ruthless" ||
    normalized === "hardcore ruthless" ||
    /^(hc|hardcore)\b/.test(normalized) ||
    /\b(ssf|ruthless)\b/.test(normalized)
  );
}

async function fetchActiveChallengeLeagues(): Promise<
  PoeTradeLeagueEntry[]
> {
  const response = await fetch(
    "https://www.pathofexile.com/api/trade2/data/leagues",
    {
      headers: {
        Accept: "application/json",
        "User-Agent": "PoE2-Collector/0.1.1",
      },
    },
  );

  if (!response.ok) {
    throw new Error(
      `League lookup failed with HTTP ${response.status}.`,
    );
  }

  const data =
    (await response.json()) as PoeTradeLeagueResponse;

  console.log(
    "PoE 2 trade leagues:",
    data.result.map((league) => league.id),
  );

  const seen = new Set<string>();

  const activeLeagues =
    data.result.filter((league) => {
      if (
        isExcludedLeagueKey(
          league.id,
        )
      ) {
        return false;
      }

      if (seen.has(league.id)) {
        return false;
      }

      seen.add(league.id);
      return true;
    });

  console.log(
    "Detected active PoE 2 leagues:",
    activeLeagues.map(
      (league) => league.id,
    ),
  );

  return activeLeagues;
}
async function initializeCollectionProfiles(
  db: Database,
): Promise<{
  profiles: CollectionProfile[];
  activeProfileId: string;
}> {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS collection_profiles (
      id TEXT PRIMARY KEY NOT NULL,
      name TEXT NOT NULL,
      kind TEXT NOT NULL,
      league_key TEXT,
      is_archived INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await db.execute(`
    CREATE TABLE IF NOT EXISTS profile_unique_tracking (
      profile_id TEXT NOT NULL,
      unique_id TEXT NOT NULL,
      flag TEXT NOT NULL,
      PRIMARY KEY (profile_id, unique_id, flag)
    )
  `);

  await db.execute(`
    CREATE TABLE IF NOT EXISTS profile_collection_review (
      profile_id TEXT NOT NULL,
      unique_id TEXT NOT NULL,
      reviewed INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (profile_id, unique_id)
    )
  `);

  await db.execute(`
  CREATE TABLE IF NOT EXISTS league_rollovers (
    id TEXT PRIMARY KEY NOT NULL,
    archived_profile_id TEXT NOT NULL,
    old_league_key TEXT,
    old_league_name TEXT NOT NULL,
    new_league_key TEXT NOT NULL,
    new_league_name TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  )
`);

await db.execute(`
  CREATE TABLE IF NOT EXISTS league_rollover_changes (
    rollover_id TEXT NOT NULL,
    unique_id TEXT NOT NULL,
    flag TEXT NOT NULL,
    PRIMARY KEY (
      rollover_id,
      unique_id,
      flag
    )
  )
`);

  await db.execute(
    `
      INSERT OR IGNORE INTO collection_profiles (
        id,
        name,
        kind,
        league_key,
        is_archived,
        sort_order
      )
      VALUES (?, ?, 'standard', NULL, 0, 0)
    `,
    [STANDARD_PROFILE_ID, "Standard"],
  );

  await db.execute(
    `
      INSERT OR IGNORE INTO collection_profiles (
        id,
        name,
        kind,
        league_key,
        is_archived,
        sort_order
      )
      VALUES (?, ?, 'challenge', NULL, 0, 1)
    `,
    [CURRENT_LEAGUE_PROFILE_ID, "Current League"],
  );

  const migrationState = await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'collection_profiles_migrated_v1'
  `);

  if (migrationState.length === 0) {
    // Everything tracked before profiles existed belongs to Standard.
    // The old tables are intentionally left untouched as a safety backup.
    await db.execute(
      `
        INSERT OR IGNORE INTO profile_unique_tracking (
          profile_id,
          unique_id,
          flag
        )
        SELECT ?, unique_id, flag
        FROM unique_tracking
      `,
      [STANDARD_PROFILE_ID],
    );

    await db.execute(
      `
        INSERT OR IGNORE INTO profile_collection_review (
          profile_id,
          unique_id,
          reviewed
        )
        SELECT ?, unique_id, reviewed
        FROM collection_review
      `,
      [STANDARD_PROFILE_ID],
    );

    // A newly created challenge league is a fresh collection: known
    // catalogue entries begin as Missing, not Unreviewed. New catalogue
    // additions discovered later can still arrive as Unreviewed.
    await db.execute(
      `
        INSERT OR IGNORE INTO profile_collection_review (
          profile_id,
          unique_id,
          reviewed
        )
        SELECT ?, id, 1
        FROM unique_variants
        WHERE source IN ('poewiki', 'built-in-special')
      `,
      [CURRENT_LEAGUE_PROFILE_ID],
    );

    await db.execute(
      `
        INSERT OR REPLACE INTO app_meta (key, value)
        VALUES ('collection_profiles_migrated_v1', 'yes')
      `,
    );

    await db.execute(
      `
        INSERT OR IGNORE INTO app_meta (key, value)
        VALUES ('active_collection_profile', ?)
      `,
      [STANDARD_PROFILE_ID],
    );
  }

  const profileRows = await db.select<
    {
      id: string;
      name: string;
      kind: string;
      is_archived: number;
    }[]
  >(`
    SELECT
      id,
      name,
      kind,
      is_archived
    FROM collection_profiles
    WHERE is_archived = 0
    ORDER BY sort_order ASC, name ASC
  `);

  const profiles: CollectionProfile[] =
    profileRows
      .filter(
        (
          row,
        ): row is {
          id: string;
          name: string;
          kind: "standard" | "challenge";
          is_archived: number;
        } =>
          row.kind === "standard" ||
          row.kind === "challenge",
      )
      .map((row) => ({
        id: row.id,
        name: row.name,
        kind: row.kind,
        isArchived: row.is_archived === 1,
      }));

  const activeRows = await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'active_collection_profile'
  `);

  const requestedActive =
    activeRows[0]?.value ?? STANDARD_PROFILE_ID;

  const activeProfileId = profiles.some(
    (profile) => profile.id === requestedActive,
  )
    ? requestedActive
    : STANDARD_PROFILE_ID;

  if (activeProfileId !== requestedActive) {
    await db.execute(
      `
        INSERT OR REPLACE INTO app_meta (key, value)
        VALUES ('active_collection_profile', ?)
      `,
      [activeProfileId],
    );
  }

  return {
    profiles,
    activeProfileId,
  };
}

async function seedBuiltInSpecialVariants(db: Database) {
  // PoE 2 currently has no hand-written special variant rules,
  // but the parser still expects this table to exist.
  await db.execute(`
    CREATE TABLE IF NOT EXISTS unique_variant_rules (
      id TEXT PRIMARY KEY NOT NULL,
      variant_id TEXT NOT NULL,
      match_type TEXT NOT NULL,
      match_text TEXT NOT NULL,
      priority INTEGER NOT NULL DEFAULT 100,
      source TEXT NOT NULL DEFAULT 'built-in'
    )
  `);

  return 0;
}

const DEFAULT_POE_LOOKUP_HOTKEY =
  "CommandOrControl+Alt+C";

const OLD_DEFAULT_POE_LOOKUP_HOTKEY =
  "CommandOrControl+Shift+C";

function formatHotkeyForDisplay(
  hotkey: string,
) {
  return hotkey
    .replace(
      "CommandOrControl",
      "Ctrl",
    )
    .split("+")
    .join(" + ");
}

function normalizeCapturedHotkeyKey(
  key: string,
) {
  const upper =
    key.toUpperCase();

  if (/^[A-Z0-9]$/.test(upper)) {
    return upper;
  }

  if (
    /^F([1-9]|1[0-2])$/.test(
      upper,
    )
  ) {
    return upper;
  }

  return null;
}

const DEV_LEAGUE_ROLLOVER_HOTKEY =
  "CommandOrControl+Shift+F12";

const OVERLAY_LABEL = "poe-overlay";
const OVERLAY_WIDTH = 430;
const OVERLAY_HEIGHT = 330;

type OverlayProfileState = {
  id: string;
  name: string;
  flags: TrackingFlag[];
};

type OverlayPayload = {
  trackingProfileId: string;
  trackingProfileName: string;
  uniqueId: string;
  name: string;
  baseType: string | null;
  itemType: string;
  variantLabel: string | null;
  edition: "normal" | "foulborn" | "vestigial";
  profileStates: OverlayProfileState[];
  trackingFlags: TrackingFlag[];
  trackingReviewed: boolean;
  editionAvailability: EditionAvailabilityMap;
  editionSources: EditionSourceMap;
  isLegacyOnly: boolean;
  statusColors: StatusColors;
  collectionRules: CollectionRules;
  extraTracking: ExtraTracking;
};

type OverlayAction =
  | {
      kind: "toggle";
      profileId: string;
      uniqueId: string;
      flag: TrackingFlag;
      enabled: boolean;
    }
  | {
      kind: "missing";
      profileId: string;
      uniqueId: string;
    };

function getOverlayPayloadFromUrl(): OverlayPayload | null {
  const params = new URLSearchParams(window.location.search);
  const encoded = params.get("data");

  if (!encoded) {
    return null;
  }

  try {
    return JSON.parse(
      decodeURIComponent(encoded),
    ) as OverlayPayload;
  } catch (error) {
    console.error("Could not parse overlay payload:", error);
    return null;
  }
}

function OverlayApp() {
  const initialPayload = useMemo(
    () => getOverlayPayloadFromUrl(),
    [],
  );

  const [payload, setPayload] =
  useState<OverlayPayload | null>(initialPayload);

const [overlayMessage, setOverlayMessage] =
  useState("Checking hovered item...");

  const lastMatchedNameRef =
  useRef<string | null>(
    initialPayload?.name ?? null,
  );

const suppressNextBlurRef =
  useRef(false);

function hideOverlayManually() {
  suppressNextBlurRef.current = true;

  void WebviewWindow.getCurrent().hide();
}

useEffect(() => {
  document.documentElement.classList.add(
    "poe-overlay-document",
  );
  document.body.classList.add(
    "poe-overlay-document",
  );

  const currentWindow =
    WebviewWindow.getCurrent();

  let focusCheckTimer:
  | number
  | undefined;

const resolveFocusChange =
  async () => {
    try {
      const mainWindow =
        await WebviewWindow.getByLabel(
          "main",
        );

      if (
        mainWindow &&
        await mainWindow.isFocused()
      ) {
        const itemName =
          lastMatchedNameRef.current;

        if (itemName) {
          await emit<string>(
            "poe-overlay-search-item",
            itemName,
          );
        }

        await currentWindow.hide();
        return;
      }

      const foregroundProcess =
        await invoke<string>(
          "get_foreground_process_name",
        );

      const normalizedProcess =
        foregroundProcess.toLowerCase();

      const isSnippingTool =
        normalizedProcess ===
          "snippingtool.exe" ||
        normalizedProcess ===
          "screenclippinghost.exe";

      if (isSnippingTool) {
        focusCheckTimer =
          window.setTimeout(
            () => {
              void resolveFocusChange();
            },
            250,
          );

        return;
      }

      await currentWindow.hide();
    } catch (error) {
      console.error(
        "Could not resolve overlay focus change:",
        error,
      );

      void currentWindow.hide();
    }
  };

const handleBlur = () => {
  if (suppressNextBlurRef.current) {
    suppressNextBlurRef.current = false;
    return;
  }

  if (focusCheckTimer !== undefined) {
    window.clearTimeout(
      focusCheckTimer,
    );
  }

  focusCheckTimer =
    window.setTimeout(
      () => {
        void resolveFocusChange();
      },
      150,
    );
};

  let unlistenLoading:
    | (() => void)
    | undefined;

  let unlistenResult:
    | (() => void)
    | undefined;

  let unlistenMessage:
    | (() => void)
    | undefined;

  window.addEventListener("blur", handleBlur);

  void listen<void>(
    "poe-overlay-loading",
    () => {
      lastMatchedNameRef.current = null;
      setPayload(null);
      setOverlayMessage(
        "Checking hovered item...",
      );
    },
  ).then((unlisten) => {
    unlistenLoading = unlisten;
  });

  void listen<OverlayPayload>(
    "poe-overlay-result",
    (event) => {
      lastMatchedNameRef.current =
        event.payload.name;

      setPayload(event.payload);
    },
  ).then((unlisten) => {
    unlistenResult = unlisten;
  });

  void listen<string>(
    "poe-overlay-message",
    (event) => {
      lastMatchedNameRef.current = null;
      setPayload(null);
      setOverlayMessage(event.payload);
    },
  ).then((unlisten) => {
    unlistenMessage = unlisten;
  });


  return () => {
    window.removeEventListener(
      "blur",
      handleBlur,
    );

    if (focusCheckTimer !== undefined) {
      window.clearTimeout(
        focusCheckTimer,
      );
    }

    unlistenLoading?.();
    unlistenResult?.();
    unlistenMessage?.();

    document.documentElement.classList.remove(
      "poe-overlay-document",
    );
    document.body.classList.remove(
      "poe-overlay-document",
    );
  };
}, []);

if (!payload) {
  return (
    <div className="poe-overlay-card overlay-error-card">
      <div className="poe-overlay-header">
        <div className="poe-overlay-title-wrap">
          <span className="poe-overlay-kicker">
            PoE 2 Unique Tracker
          </span>
          <h1>Unique Tracker</h1>
        </div>

        <button
          type="button"
          className="poe-overlay-close"
          aria-label="Close overlay"
          onClick={hideOverlayManually}
        >
          {"\u00D7"}
        </button>
      </div>

      <span>{overlayMessage}</span>
    </div>
  );
}

  const trackingIsMissing =
  payload.trackingReviewed &&
  !payload.trackingFlags.some(
    (flag) => payload.collectionRules[flag],
  );

const editionLabel =
  payload.edition === "foulborn"
    ? "FOULBORN"
    : payload.edition === "vestigial"
      ? "VESTIGIAL"
      : null;

const wornProfiles =
  payload.profileStates.filter(
    (profile) =>
      profile.flags.includes("wearing"),
  );

const ownedProfiles =
  payload.profileStates.filter(
    (profile) =>
      profile.flags.length > 0 &&
      !profile.flags.includes("wearing"),
  );

const statusHeading =
  "COLLECTION STATUS";

const statusValue =
  wornProfiles.length === 0 &&
  ownedProfiles.length === 0
    ? "Unowned"
    : [
        wornProfiles.length > 0
          ? `Currently worn in ${wornProfiles
              .map((profile) => profile.name)
              .join(", ")}`
          : null,
        ownedProfiles.length > 0
          ? `Owned in ${ownedProfiles
              .map((profile) => profile.name)
              .join(", ")}`
          : null,
      ]
        .filter(
          (value): value is string =>
            value !== null,
        )
        .join(" • ");

  const trackingHeading =
    payload.trackingProfileId === STANDARD_PROFILE_ID
      ? "Update Standard"
      : `Record in ${payload.trackingProfileName}`;

      const showFoil =
  shouldShowEdition(
    "foil",
    payload.editionAvailability,
    payload.trackingFlags,
    payload.extraTracking,
    payload.edition,
  );

const showFoulborn =
  shouldShowEdition(
    "foulborn",
    payload.editionAvailability,
    payload.trackingFlags,
    payload.extraTracking,
    payload.edition,
  );

const showVestigial =
  shouldShowEdition(
    "vestigial",
    payload.editionAvailability,
    payload.trackingFlags,
    payload.extraTracking,
    payload.edition,
  );

  const foilUncertain =
  showFoil &&
  isEditionUncertain(
    "foil",
    payload.editionAvailability,
    payload.editionSources,
    payload.trackingFlags,
    payload.edition,
  );

const foulbornUncertain =
  showFoulborn &&
  isEditionUncertain(
    "foulborn",
    payload.editionAvailability,
    payload.editionSources,
    payload.trackingFlags,
    payload.edition,
  );

const vestigialUncertain =
  showVestigial &&
  isEditionUncertain(
    "vestigial",
    payload.editionAvailability,
    payload.editionSources,
    payload.trackingFlags,
    payload.edition,
  );

const hasUncertainEdition =
  foilUncertain ||
  foulbornUncertain ||
  vestigialUncertain;

  async function toggleFlag(flag: TrackingFlag) {
    if (!payload) {
      return;
    }

    const enabled =
      !payload.trackingFlags.includes(flag);

    const enabledFlags: TrackingFlag[] =
      enabled && flag === "foil"
        ? ["owned", "foil"]
        : [flag];

    setPayload((current) => {
  if (!current) {
    return current;
  }

  const nextTrackingFlags =
    enabled
      ? Array.from(
          new Set([
            ...current.trackingFlags,
            ...enabledFlags,
          ]),
        )
      : current.trackingFlags.filter(
          (existingFlag) =>
            existingFlag !== flag,
        );

  return {
    ...current,
    trackingReviewed: true,
    trackingFlags:
      nextTrackingFlags,
    profileStates:
      current.profileStates.map(
        (profile) =>
          profile.id ===
          current.trackingProfileId
            ? {
                ...profile,
                flags:
                  nextTrackingFlags,
              }
            : profile,
      ),
  };
});

    await emit<OverlayAction>(
      "poe-overlay-action",
      {
        kind: "toggle",
        profileId:
          payload.trackingProfileId,
        uniqueId: payload.uniqueId,
        flag,
        enabled,
      },
    );
  }

  async function markMissing() {
    if (!payload) {
      return;
    }

    setPayload((current) =>
  current
    ? {
        ...current,
        trackingReviewed: true,
        trackingFlags: [],
        profileStates:
          current.profileStates.map(
            (profile) =>
              profile.id ===
              current.trackingProfileId
                ? {
                    ...profile,
                    flags: [],
                  }
                : profile,
          ),
      }
    : current,
);

    await emit<OverlayAction>(
      "poe-overlay-action",
      {
        kind: "missing",
        profileId: payload.trackingProfileId,
        uniqueId: payload.uniqueId,
      },
    );
  }

  return (
    <div className="poe-overlay-card">
      <div className="poe-overlay-header">
        <div className="poe-overlay-title-wrap">
          <span className="poe-overlay-kicker">
            PoE 2 Unique Tracker
          </span>
          <h1>{payload.name}</h1>

          <div className="poe-overlay-meta">
            <span>
              {payload.baseType ?? payload.itemType}
            </span>

            {payload.variantLabel && (
              <span>{payload.variantLabel}</span>
            )}
          </div>
        </div>

        <button
          type="button"
          className="poe-overlay-close"
          aria-label="Close overlay"
          onClick={hideOverlayManually}
        >
          {"\u00D7"}
        </button>
      </div>

      <div className="poe-overlay-tags">
        {editionLabel && (
          <span
            className={`poe-overlay-edition ${payload.edition}`}
          >
            {editionLabel}
          </span>
        )}

        {payload.isLegacyOnly && (
          <span className="poe-overlay-legacy">
            LEGACY
          </span>
        )}

        {!payload.trackingReviewed && (
  <span className="poe-overlay-unreviewed">
    UNREVIEWED
  </span>
)}
      </div>

      <div className="poe-overlay-status-heading">
        <span>{statusHeading}</span>
        <strong>{statusValue}</strong>
      </div>

      <div className="poe-overlay-tracking-heading">
        {trackingHeading}
      </div>

      <div className="poe-overlay-actions">
        <button
          type="button"
          className={`tracking-badge missing ${
            trackingIsMissing ? "active" : ""
          }`}
          style={getStatusStyle(
            payload.statusColors.missing,
          )}
          onClick={() => void markMissing()}
        >
          Missing
        </button>

        <button
          type="button"
          className={`tracking-badge owned ${
            payload.trackingFlags.includes("owned")
              ? "active"
              : ""
          }`}
          style={getStatusStyle(
            payload.statusColors.owned,
          )}
          onClick={() => void toggleFlag("owned")}
        >
          Normal
        </button>

        <button
          type="button"
          className={`tracking-badge wearing ${
            payload.trackingFlags.includes("wearing")
              ? "active"
              : ""
          }`}
          style={getStatusStyle(
            payload.statusColors.wearing,
          )}
          onClick={() => void toggleFlag("wearing")}
        >
          Wearing
        </button>

        {showFoil && (
  <button
    type="button"
    className={`tracking-badge foil ${
      payload.trackingFlags.includes("foil")
        ? "active"
        : ""
    } ${
      payload.statusColors.foil === "rainbow"
        ? "rainbow-status"
        : ""
    }`}
    style={getStatusStyle(
      payload.statusColors.foil,
    )}
    onClick={() => void toggleFlag("foil")}
  >
    {foilUncertain ? "Foil*" : "Foil"}
  </button>
)}

        {showFoulborn && (
  <button
    type="button"
    className={`tracking-badge foulborn ${
      payload.trackingFlags.includes("foulborn")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      payload.statusColors.foulborn,
    )}
    onClick={() => void toggleFlag("foulborn")}
  >
    {foulbornUncertain
      ? "Foulborn*"
      : "Foulborn"}
  </button>
)}

        {showVestigial && (
  <button
    type="button"
    className={`tracking-badge vestigial ${
      payload.trackingFlags.includes("vestigial")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      payload.statusColors.vestigial,
    )}
    onClick={() => void toggleFlag("vestigial")}
  >
    {vestigialUncertain
      ? "Vestigial*"
      : "Vestigial"}
  </button>
)}
      </div>

      <p
  className={`poe-overlay-hint ${
    hasUncertainEdition
      ? "uncertain"
      : ""
  }`}
>
  {hasUncertainEdition
    ? "*Data inconclusive. This edition may or may not exist."
    : "Click a status to update your collection."}
</p>
    </div>
  );
}


function wait(milliseconds: number) {
  return new Promise<void>((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

function isSqliteLockedError(
  error: unknown,
) {
  const message =
    error instanceof Error
      ? error.message
      : String(error);

  return (
    message
      .toLowerCase()
      .includes("database is locked") ||
    message.includes("(code: 5)")
  );
}

async function withSqliteLockRetry<T>(
  operation: () => Promise<T>,
): Promise<T> {
  const attempts = 4;

  for (
    let attempt = 1;
    attempt <= attempts;
    attempt += 1
  ) {
    try {
      return await operation();
    } catch (error) {
      if (
        !isSqliteLockedError(error) ||
        attempt === attempts
      ) {
        throw error;
      }

      await wait(
        150 * attempt,
      );
    }
  }

  throw new Error(
    "SQLite retry unexpectedly ended.",
  );
}

function TrackingBadges({
  flags,
  colors,
  extraTracking,
  editionAvailability,
  editionSources,
  isCollected,
  isMissing,
  onToggle,
  onMarkMissing,
}: {
  flags: TrackingFlag[];
  colors: StatusColors;
  extraTracking: ExtraTracking;
  editionAvailability: EditionAvailabilityMap;
  editionSources: EditionSourceMap;
  isCollected: boolean;
  isMissing: boolean;
  onToggle: (flag: TrackingFlag) => void;
  onMarkMissing: () => void;
}) {
  const showFoil =
    shouldShowEdition(
      "foil",
      editionAvailability,
      flags,
      extraTracking,
    );

  const showFoulborn =
    shouldShowEdition(
      "foulborn",
      editionAvailability,
      flags,
      extraTracking,
    );

  const showVestigial =
    shouldShowEdition(
      "vestigial",
      editionAvailability,
      flags,
      extraTracking,
    );

    const foilUncertain =
  showFoil &&
  isEditionUncertain(
    "foil",
    editionAvailability,
    editionSources,
    flags,
  );

const foulbornUncertain =
  showFoulborn &&
  isEditionUncertain(
    "foulborn",
    editionAvailability,
    editionSources,
    flags,
  );

const vestigialUncertain =
  showVestigial &&
  isEditionUncertain(
    "vestigial",
    editionAvailability,
    editionSources,
    flags,
  );

  const hasAnyExtraTracking =
    showFoil ||
    showFoulborn ||
    showVestigial;

  return (
    <div className="tracking-badges">
      <button
        type="button"
        className={`tracking-badge missing ${isMissing ? "active" : ""}`}
        style={getStatusStyle(colors.missing)}
        onClick={onMarkMissing}
      >
        Missing
      </button>

      <button
        type="button"
        className={`tracking-badge owned ${isCollected ? "active" : ""}`}
        style={getStatusStyle(colors.owned)}
        onClick={() => onToggle("owned")}
      >
        Owned
      </button>

      <button
        type="button"
        className={`tracking-badge wearing ${flags.includes("wearing") ? "active" : ""}`}
        style={getStatusStyle(colors.wearing)}
        onClick={() => onToggle("wearing")}
      >
        Wearing
      </button>

      {hasAnyExtraTracking && (
        <span className="tracking-divider" />
      )}

      {showFoil && (
  <button
    type="button"
    className={`tracking-badge foil ${
      flags.includes("foil") ? "active" : ""
    } ${
      colors.foil === "rainbow"
        ? "rainbow-status"
        : ""
    }`}
    style={getStatusStyle(colors.foil)}
    onClick={() => onToggle("foil")}
  >
    {foilUncertain ? "Foil*" : "Foil"}
  </button>
)}

      {showFoulborn && (
  <button
    type="button"
    className={`tracking-badge foulborn ${
      flags.includes("foulborn")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      colors.foulborn,
    )}
    onClick={() =>
      onToggle("foulborn")
    }
  >
    {foulbornUncertain
      ? "Foulborn*"
      : "Foulborn"}
  </button>
)}

      {showVestigial && (
  <button
    type="button"
    className={`tracking-badge vestigial ${
      flags.includes("vestigial")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      colors.vestigial,
    )}
    onClick={() =>
      onToggle("vestigial")
    }
  >
    {vestigialUncertain
      ? "Vestigial*"
      : "Vestigial"}
  </button>
)}
    </div>
  );
}

function MainApp() {
  const [uniques, setUniques] = useState<UniqueEntry[]>([]);
  const [sourceFile, setSourceFile] = useState("");
  const [sourceImportMode, setSourceImportMode] =
    useState<ImportMode | null>(null);
  const [appError, setAppError] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [database, setDatabase] = useState<Database | null>(null);
  const [databaseReady, setDatabaseReady] = useState(false);
  const [
  closePromptOpen,
  setClosePromptOpen,
] = useState(false);
  const [collectionProfiles, setCollectionProfiles] = useState<
    CollectionProfile[]
  >([]);
  const [activeProfileId, setActiveProfileId] = useState(
    STANDARD_PROFILE_ID,
  );
  const hotkeyLookupRef = useRef<
    () => Promise<void>
  >(async () => {});

  const registeredLookupHotkeyRef =
    useRef<string | null>(null);

  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("all");
  const [typeFilter, setTypeFilter] = useState("All");
  const [sortMode, setSortMode] = useState<SortMode>("alphabetical");

  const [settingsOpen, setSettingsOpen] = useState(false);
  const [statusColors, setStatusColors] =
    useState<StatusColors>(DEFAULT_STATUS_COLORS);
  const [collectionRules, setCollectionRules] =
    useState<CollectionRules>(DEFAULT_COLLECTION_RULES);
  const [extraTracking, setExtraTracking] =
    useState<ExtraTracking>(DEFAULT_EXTRA_TRACKING);
  const [showLegacyVariants, setShowLegacyVariants] =
    useState(true);

  const [catalogueUpdate, setCatalogueUpdate] =
    useState<CatalogueUpdateSummary | null>(null);
  const [catalogueUpdateOpen, setCatalogueUpdateOpen] =
    useState(false);
  const [catalogueChecking, setCatalogueChecking] =
    useState(false);
  const [catalogueCheckMessage, setCatalogueCheckMessage] =
    useState("Catalogue updates have not been checked yet.");

  const [
    discoveryNotices,
    setDiscoveryNotices,
  ] = useState<TrackerDiscoveryNotice[]>([]);
  const [
    discoveryDetailsExpanded,
    setDiscoveryDetailsExpanded,
  ] = useState(false);
  const [
    devDiscoveryNotice,
    setDevDiscoveryNotice,
  ] = useState<TrackerDiscoveryNotice | null>(
    null,
  );

  const [importMatchSummary, setImportMatchSummary] =
    useState<ImportReconciliationSummary | null>(null);

  const [pendingImport, setPendingImport] =
    useState<PendingImport | null>(null);
  const [pendingImportMode, setPendingImportMode] =
    useState<ImportMode>("status-list");
  const [
    otherImportModesOpen,
    setOtherImportModesOpen,
  ] = useState(false);
  const [pendingImportProfileId, setPendingImportProfileId] =
    useState(STANDARD_PROFILE_ID);
  const [missingOnlySummary, setMissingOnlySummary] =
    useState<MissingOnlyInferenceSummary | null>(null);
  const [colorCodedSummary, setColorCodedSummary] =
    useState<ColorCodedImportSummary | null>(null);
  const [importDetailsExpanded, setImportDetailsExpanded] =
    useState(false);

  const [
    batchImportReviewOpen,
    setBatchImportReviewOpen,
  ] = useState(false);

  const [parserTestText, setParserTestText] = useState("");
  const [parserTestResult, setParserTestResult] =
    useState<ItemIdentificationResult | null>(null);
  const [parserTesting, setParserTesting] = useState(false);
  const [
  lookupHotkey,
  setLookupHotkey,
] = useState(
  DEFAULT_POE_LOOKUP_HOTKEY,
);

const [
  hotkeyRecording,
  setHotkeyRecording,
] = useState(false);

const [
  hotkeySettingMessage,
  setHotkeySettingMessage,
] = useState("");
  const [hotkeyReady, setHotkeyReady] = useState(false);
  const [hotkeyBusy, setHotkeyBusy] = useState(false);
  const [hotkeyMessage, setHotkeyMessage] = useState(
    "Global hotkey is starting...",
  );

  const [rolloverPreview, setRolloverPreview] =
  useState<LeagueRolloverPreview | null>(null);

const [rolloverPreviewOpen, setRolloverPreviewOpen] =
  useState(false);

  const [rolloverMode, setRolloverMode] =
  useState<LeagueRolloverMode>(
    "dev-pending",
  );

const [rolloverApplying, setRolloverApplying] =
  useState(false);

const [
  rolloverChangesExpanded,
  setRolloverChangesExpanded,
] = useState(false);

  function isUniqueCollected(unique: UniqueEntry) {
    return (
      unique.reviewed &&
      unique.flags.some(
        (flag) => collectionRules[flag],
      )
    );
  }

  function isUniqueMissing(unique: UniqueEntry) {
    return (
      unique.reviewed &&
      !isUniqueCollected(unique)
    );
  }

  const activeProfile = useMemo(
  () =>
    collectionProfiles.find(
      (profile) => profile.id === activeProfileId,
    ) ?? null,
  [collectionProfiles, activeProfileId],
);

useEffect(() => {
  let unlistenClose:
    | (() => void)
    | undefined;

  void listen<void>(
    "main-close-requested",
    () => {
      setClosePromptOpen(true);
    },
  ).then((unlisten) => {
    unlistenClose = unlisten;
  });

  return () => {
    unlistenClose?.();
  };
}, []);

async function initializeLiveLeagues(
    db: Database,
    profiles: CollectionProfile[],
  ) {
    const newlyDetectedLeagues:
      PendingLeagueDetection[] = [];

    try {
      const detectedLeagues =
        await fetchActiveChallengeLeagues();

      const activeLeagueKeys =
        new Set(
          detectedLeagues.map(
            (league) => league.id,
          ),
        );

      const storedRows =
        await db.select<
          {
            id: string;
            name: string;
            league_key: string | null;
            created_at: string;
          }[]
        >(`
          SELECT
            id,
            name,
            league_key,
            created_at
          FROM collection_profiles
          WHERE
            kind = 'challenge'
            AND is_archived = 0
          ORDER BY
            sort_order ASC,
            created_at ASC
        `);

      const removableExcludedIds =
        new Set<string>();

      for (const profile of storedRows) {
        if (
          !profile.league_key ||
          !isExcludedLeagueKey(
            profile.league_key,
          )
        ) {
          continue;
        }

        const trackingRows =
          await db.select<
            { total: number }[]
          >(
            `
              SELECT COUNT(*) AS total
              FROM profile_unique_tracking
              WHERE profile_id = ?
            `,
            [profile.id],
          );

        const trackedCount =
          Number(
            trackingRows[0]?.total ?? 0,
          );

        if (trackedCount > 0) {
          console.warn(
            `Excluded league profile ${profile.name} contains tracked data and was kept for safety.`,
          );
          continue;
        }

        await db.execute(
          `
            DELETE FROM profile_collection_review
            WHERE profile_id = ?
          `,
          [profile.id],
        );

        await db.execute(
          `
            DELETE FROM collection_profiles
            WHERE id = ?
          `,
          [profile.id],
        );

        removableExcludedIds.add(
          profile.id,
        );
      }

      const storedProfiles =
        storedRows.filter(
          (profile) =>
            !removableExcludedIds.has(
              profile.id,
            ),
        );

      let nextProfiles =
        profiles.filter(
          (profile) =>
            !removableExcludedIds.has(
              profile.id,
            ),
        );

      /*
       * Older databases have one special "current-league"
       * row. If it has never been assigned a live league,
       * adopt one active league without touching its data.
       *
       * Existing databases that already say Runes of Aldur
       * keep that exact association.
       */
      const legacyUnassigned =
        storedProfiles.find(
          (profile) =>
            profile.id ===
              CURRENT_LEAGUE_PROFILE_ID &&
            !profile.league_key,
        );

      if (
        legacyUnassigned &&
        detectedLeagues.length > 0
      ) {
        const adoptedLeague =
          detectedLeagues.find(
            (league) =>
              league.id ===
              legacyUnassigned.name,
          ) ??
          detectedLeagues[0];

        await db.execute(
          `
            UPDATE collection_profiles
            SET
              name = ?,
              league_key = ?
            WHERE id = ?
          `,
          [
            adoptedLeague.id,
            adoptedLeague.id,
            legacyUnassigned.id,
          ],
        );

        legacyUnassigned.name =
          adoptedLeague.id;

        legacyUnassigned.league_key =
          adoptedLeague.id;

        nextProfiles =
          nextProfiles.map(
            (profile) =>
              profile.id ===
              legacyUnassigned.id
                ? {
                    ...profile,
                    name:
                      adoptedLeague.id,
                  }
                : profile,
          );
      }

      const representedLeagueKeys =
        new Set(
          storedProfiles
            .map(
              (profile) =>
                profile.league_key,
            )
            .filter(
              (
                leagueKey,
              ): leagueKey is string =>
                Boolean(leagueKey),
            ),
        );

      let nextSortOrder = 10;

      /*
       * Every newly detected active event/league gets its
       * own fresh collection. Existing active profiles are
       * never replaced merely because another league appears.
       */
      for (
        const league of
        detectedLeagues
      ) {
        if (
          representedLeagueKeys.has(
            league.id,
          )
        ) {
          continue;
        }

        const baseProfileId =
          makeLeagueProfileId(
            league.id,
          );

        const existingIdRows =
          await db.select<
            { id: string }[]
          >(
            `
              SELECT id
              FROM collection_profiles
              WHERE id = ?
            `,
            [baseProfileId],
          );

        const profileId =
          existingIdRows.length === 0
            ? baseProfileId
            : `${baseProfileId}:${Date.now()}:${nextSortOrder}`;

        await db.execute(
          `
            INSERT INTO collection_profiles (
              id,
              name,
              kind,
              league_key,
              is_archived,
              sort_order
            )
            VALUES (
              ?,
              ?,
              'challenge',
              ?,
              0,
              ?
            )
          `,
          [
            profileId,
            league.id,
            league.id,
            String(nextSortOrder),
          ],
        );

        /*
         * A new active league starts as a clean collection.
         * Existing catalogue entries are known Missing,
         * rather than appearing Unreviewed.
         */
        await db.execute(
          `
            INSERT OR IGNORE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT
              ?,
              id,
              1
            FROM unique_variants
            WHERE source IN (
              'poewiki',
              'built-in-special'
            )
          `,
          [profileId],
        );

        storedProfiles.push({
          id: profileId,
          name: league.id,
          league_key: league.id,
          created_at:
            new Date().toISOString(),
        });

        representedLeagueKeys.add(
          league.id,
        );

        nextProfiles.push({
          id: profileId,
          name: league.id,
          kind: "challenge",
          isArchived: false,
        });

        const newLeagueDetection = {
          profileId,
          leagueName: league.id,
        };

        newlyDetectedLeagues.push(
          newLeagueDetection,
        );

        /*
         * Save the pending announcement immediately. If a later
         * league-startup step or Wiki request fails, the next launch
         * can still finish the refresh and show the notice.
         */
        await appendPendingLeagueDetections(
          db,
          [newLeagueDetection],
        );

        nextSortOrder += 1;
      }

      /*
       * A league is considered potentially ended only when
       * it used to be active locally but no longer appears
       * in GGG's current active league list.
       *
       * We only SHOW a migration prompt here. No collection
       * data moves until the user confirms it.
       */
      const endedProfiles =
        storedProfiles.filter(
          (profile) =>
            profile.league_key !== null &&
            !activeLeagueKeys.has(
              profile.league_key,
            ),
        );

      if (endedProfiles.length > 0) {
        const preview =
          await buildLeagueRolloverPreview(
            db,
            endedProfiles[0].id,
            "Standard",
          );

        setRolloverPreview(preview);
        setRolloverMode("pending");
        setRolloverChangesExpanded(false);
        setRolloverPreviewOpen(true);
      }

      return [...nextProfiles].sort(
        (left, right) => {
          if (
            left.kind === "standard" &&
            right.kind !== "standard"
          ) {
            return -1;
          }

          if (
            right.kind === "standard" &&
            left.kind !== "standard"
          ) {
            return 1;
          }

          return left.name.localeCompare(
            right.name,
          );
        },
      );
    } catch (error) {
      console.error(
        "Could not detect active Path of Exile 2 leagues:",
        error,
      );

      // League detection is optional. The local collection
      // still starts normally while offline.
      return profiles;
    }
  }
async function buildLeagueRolloverPreview(
  db: Database,
  profileId: string,
  newLeagueName: string,
): Promise<LeagueRolloverPreview> {
  const leagueRows =
    await db.select<
      {
        unique_id: string;
        flag: string;
        name: string | null;
      }[]
    >(
      `
        SELECT
          tracking.unique_id,
          tracking.flag,
          variants.name
        FROM profile_unique_tracking tracking
        LEFT JOIN unique_variants variants
          ON variants.id = tracking.unique_id
        WHERE tracking.profile_id = ?
      `,
      [profileId],
    );

  const standardRows =
    await db.select<
      {
        unique_id: string;
        flag: string;
      }[]
    >(
      `
        SELECT
          unique_id,
          flag
        FROM profile_unique_tracking
        WHERE profile_id = ?
      `,
      [STANDARD_PROFILE_ID],
    );

  const currentLeagueRows =
    await db.select<
      {
        name: string;
        league_key: string | null;
        kind: string;
      }[]
    >(
      `
        SELECT
          name,
          league_key,
          kind
        FROM collection_profiles
        WHERE id = ?
      `,
      [profileId],
    );

  const currentLeague =
    currentLeagueRows[0];

  if (
    !currentLeague ||
    currentLeague.kind !== "challenge"
  ) {
    throw new Error(
      "The league collection could not be found.",
    );
  }

  const standardFlags =
    new Set(
      standardRows.map(
        (row) =>
          `${row.unique_id}|||${row.flag}`,
      ),
    );

  const changesByUnique =
    new Map<
      string,
      LeagueRolloverChange
    >();

  const flagCounts: Record<
    TrackingFlag,
    number
  > = {
    owned: 0,
    wearing: 0,
    foil: 0,
    foulborn: 0,
    vestigial: 0,
  };

  for (const row of leagueRows) {
    if (!isTrackingFlag(row.flag)) {
      continue;
    }

    const alreadyInStandard =
      standardFlags.has(
        `${row.unique_id}|||${row.flag}`,
      );

    if (alreadyInStandard) {
      continue;
    }

    flagCounts[row.flag] += 1;

    const existing =
      changesByUnique.get(
        row.unique_id,
      );

    if (existing) {
      if (
        !existing.addedFlags.includes(
          row.flag,
        )
      ) {
        existing.addedFlags.push(
          row.flag,
        );
      }

      continue;
    }

    changesByUnique.set(
      row.unique_id,
      {
        uniqueId: row.unique_id,
        name:
          row.name ??
          row.unique_id,
        addedFlags: [row.flag],
      },
    );
  }

  const changes =
    Array.from(
      changesByUnique.values(),
    ).sort((left, right) =>
      left.name.localeCompare(
        right.name,
      ),
    );

  return {
    profileId,
    oldLeagueName:
      currentLeague.name,
    newLeagueName,
    changedUniques:
      changes.length,
    flagCounts,
    changes,
  };
}
async function executeLeagueRollover(
  db: Database,
  profileId: string,
  commitChanges: boolean,
): Promise<LeagueRolloverPreview> {
  const preview =
    await buildLeagueRolloverPreview(
      db,
      profileId,
      "Standard",
    );

  const currentLeagueRows =
    await db.select<
      {
        name: string;
        league_key: string | null;
        kind: string;
      }[]
    >(
      `
        SELECT
          name,
          league_key,
          kind
        FROM collection_profiles
        WHERE id = ?
      `,
      [profileId],
    );

  const currentLeague =
    currentLeagueRows[0];

  if (
    !currentLeague ||
    currentLeague.kind !== "challenge"
  ) {
    throw new Error(
      "The league collection could not be found.",
    );
  }

  const oldLeagueKey =
    currentLeague.league_key ??
    currentLeague.name;

  const timestamp = Date.now();

  const rolloverId =
    `rollover:${timestamp}`;

  const statements:
    SqliteTransactionStatement[] = [
      {
        sql: `
          INSERT OR IGNORE INTO profile_unique_tracking (
            profile_id,
            unique_id,
            flag
          )
          SELECT
            ?,
            unique_id,
            flag
          FROM profile_unique_tracking
          WHERE profile_id = ?
        `,
        params: [
          STANDARD_PROFILE_ID,
          profileId,
        ],
      },
      {
        sql: `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          SELECT
            ?,
            unique_id,
            1
          FROM profile_unique_tracking
          WHERE profile_id = ?
          GROUP BY unique_id
        `,
        params: [
          STANDARD_PROFILE_ID,
          profileId,
        ],
      },
      {
        sql: `
          INSERT INTO league_rollovers (
            id,
            archived_profile_id,
            old_league_key,
            old_league_name,
            new_league_key,
            new_league_name
          )
          VALUES (?, ?, ?, ?, ?, ?)
        `,
        params: [
          rolloverId,
          profileId,
          oldLeagueKey,
          currentLeague.name,
          STANDARD_PROFILE_ID,
          "Standard",
        ],
      },
    ];

  for (const change of preview.changes) {
    for (
      const flag of
      change.addedFlags
    ) {
      statements.push({
        sql: `
          INSERT INTO league_rollover_changes (
            rollover_id,
            unique_id,
            flag
          )
          VALUES (?, ?, ?)
        `,
        params: [
          rolloverId,
          change.uniqueId,
          flag,
        ],
      });
    }
  }

  statements.push({
    sql: `
      UPDATE collection_profiles
      SET
        is_archived = 1,
        sort_order = 100
      WHERE id = ?
    `,
    params: [profileId],
  });

  await invoke(
    "execute_sqlite_transaction",
    {
      statements,
      commit: commitChanges,
    },
  );

  return preview;
}
async function previewLeagueRollover() {
  if (!database) {
    return;
  }

  try {
    setAppError("");

    const testProfile =
      collectionProfiles.find(
        (profile) =>
          profile.id ===
            activeProfileId &&
          profile.kind ===
            "challenge",
      ) ??
      collectionProfiles.find(
        (profile) =>
          profile.kind ===
          "challenge",
      );

    if (!testProfile) {
      throw new Error(
        "No active league profile is available for the rollover test.",
      );
    }

    const preview =
      await buildLeagueRolloverPreview(
        database,
        testProfile.id,
        "Standard",
      );

    setRolloverPreview(preview);
    setRolloverMode("dev-pending");
    setRolloverChangesExpanded(false);
    setRolloverPreviewOpen(true);

    const mainWindow =
      await WebviewWindow.getByLabel(
        "main",
      );

    if (mainWindow) {
      await mainWindow.show();
      await mainWindow.setFocus();
    }
  } catch (error) {
    console.error(
      "Could not simulate league ending:",
      error,
    );

    setAppError(
      error instanceof Error
        ? error.message
        : String(error),
    );
  }
}

async function confirmDevLeagueRollover() {
  if (
    !database ||
    !rolloverPreview ||
    rolloverApplying
  ) {
    return;
  }

  try {
    setRolloverApplying(true);
    setAppError("");

    const completed =
      await executeLeagueRollover(
        database,
        rolloverPreview.profileId,
        false,
      );

    setRolloverPreview(completed);
    setRolloverMode("dev-complete");
    setRolloverChangesExpanded(false);
  } catch (error) {
    console.error(
      "Could not test league migration:",
      error,
    );

    setAppError(
      error instanceof Error
        ? error.message
        : String(error),
    );
  } finally {
    setRolloverApplying(false);
  }
}

async function confirmLeagueRollover() {
  if (
    !database ||
    !rolloverPreview ||
    rolloverApplying
  ) {
    return;
  }

  try {
    setRolloverApplying(true);
    setAppError("");

    const completed =
      await executeLeagueRollover(
        database,
        rolloverPreview.profileId,
        true,
      );

    const archivedWasActive =
      activeProfileId ===
      completed.profileId;

    const nextActiveProfileId =
      archivedWasActive
        ? STANDARD_PROFILE_ID
        : activeProfileId;

    setCollectionProfiles(
      (current) =>
        current.filter(
          (profile) =>
            profile.id !==
            completed.profileId,
        ),
    );

    if (archivedWasActive) {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (
            key,
            value
          )
          VALUES (
            'active_collection_profile',
            ?
          )
        `,
        [STANDARD_PROFILE_ID],
      );

      setActiveProfileId(
        STANDARD_PROFILE_ID,
      );
    }

    await loadCollectionData(
      database,
      nextActiveProfileId,
    );

    setRolloverPreview(completed);
    setRolloverMode("complete");
    setRolloverChangesExpanded(false);
  } catch (error) {
    console.error(
      "Could not complete league migration:",
      error,
    );

    setAppError(
      error instanceof Error
        ? error.message
        : String(error),
    );
  } finally {
    setRolloverApplying(false);
  }
}
async function loadImportReconciliationSummary(
  db: Database,
): Promise<ImportReconciliationSummary> {
  const importedRows = await db.select<
    { total: number }[]
  >(`
    SELECT COUNT(*) AS total
    FROM imported_collection
  `);

  const importedTotal = Number(
    importedRows[0]?.total ?? 0,
  );

  const tableRows = await db.select<
    { name: string }[]
  >(`
    SELECT name
    FROM sqlite_master
    WHERE
      type = 'table'
      AND name = 'import_reconciliation'
  `);

  if (tableRows.length === 0) {
    return reconcileImportedCollection(db);
  }

  const reconciledRows = await db.select<
    { total: number }[]
  >(`
    SELECT COUNT(*) AS total
    FROM import_reconciliation
  `);

  const reconciledTotal = Number(
    reconciledRows[0]?.total ?? 0,
  );

  if (reconciledTotal !== importedTotal) {
    return reconcileImportedCollection(db);
  }

  const statusRows = await db.select<
    {
      status: string;
      total: number;
    }[]
  >(`
    SELECT
      status,
      COUNT(*) AS total
    FROM import_reconciliation
    GROUP BY status
  `);

  const summary: ImportReconciliationSummary = {
    matched: 0,
    ambiguous: 0,
    unmatched: 0,
    skipped: 0,
    total: importedTotal,
  };

  for (const row of statusRows) {
    if (
      row.status === "matched" ||
      row.status === "ambiguous" ||
      row.status === "unmatched" ||
      row.status === "skipped"
    ) {
      summary[row.status] = Number(row.total);
    }
  }

  return summary;
}

  async function loadCollectionData(
    db: Database,
    profileId = activeProfileId,
  ) {
    const canonicalReady =
      await isCanonicalCatalogueReady(db);

    const catalogueItems = await db.select<
      {
        id: string;
        family_id: string;
        name: string;
        base_type: string | null;
        item_type: string;
        variant_label: string | null;
        is_legacy_only: number;
        source: string;
      }[]
    >(`
      SELECT
        v.id,
        v.family_id,
        v.name,
        v.base_type,
        v.item_type,
        v.variant_label,
        v.is_legacy_only,
        v.source
      FROM unique_variants v
      WHERE
        ${
          canonicalReady
            ? "v.source IN ('poewiki', 'built-in-special')"
            : "1 = 1"
        }
        AND NOT (
          v.source = 'poewiki'
          AND EXISTS (
            SELECT 1
            FROM unique_variants special
            WHERE
              special.family_id = v.family_id
              AND special.source = 'built-in-special'
          )
        )
    `);

    const importedItems = await db.select<
      {
        id: string;
        status: string;
      }[]
    >(`
      SELECT id, status
      FROM imported_collection
    `);

    const savedTracking = await db.select<
      {
        unique_id: string;
        flag: string;
      }[]
    >(
      `
        SELECT unique_id, flag
        FROM profile_unique_tracking
        WHERE profile_id = ?
      `,
      [profileId],
    );

    const savedReview = await db.select<
      {
        unique_id: string;
        reviewed: number;
      }[]
    >(
      `
        SELECT unique_id, reviewed
        FROM profile_collection_review
        WHERE profile_id = ?
      `,
      [profileId],
    );

    const editionRows =
  await db.select<
    {
      variant_id: string;
      edition: string;
      availability: string;
      source: string;
    }[]
  >(`
    SELECT
      variant_id,
      edition,
      availability,
      source
    FROM unique_variant_editions
  `);

    const importedStatusById = new Map<string, string>(
      importedItems.map((item) => [
        item.id,
        item.status,
      ] as [string, string]),
    );

    const flagsByUnique = new Map<string, TrackingFlag[]>();

    for (const row of savedTracking) {
      if (!isTrackingFlag(row.flag)) {
        continue;
      }

      const existing =
        flagsByUnique.get(row.unique_id) ?? [];

      existing.push(row.flag);
      flagsByUnique.set(
        row.unique_id,
        existing,
      );
    }

    const reviewedByUnique = new Map<string, boolean>(
      savedReview.map((row) => [
        row.unique_id,
        row.reviewed === 1,
      ] as [string, boolean]),
    );

    const editionAvailabilityByUnique =
  new Map<
    string,
    EditionAvailabilityMap
  >();

  const editionSourcesByUnique =
  new Map<
    string,
    EditionSourceMap
  >();

for (const row of editionRows) {
  if (
    !isExtraVariant(row.edition) ||
    !isEditionAvailability(
      row.availability,
    )
  ) {
    continue;
  }

  const existing =
    editionAvailabilityByUnique.get(
      row.variant_id,
    ) ?? {
      ...DEFAULT_EDITION_AVAILABILITY,
    };

  existing[row.edition] =
    row.availability;

  editionAvailabilityByUnique.set(
    row.variant_id,
    existing,
  );

  const existingSources =
  editionSourcesByUnique.get(
    row.variant_id,
  ) ?? {
    ...DEFAULT_EDITION_SOURCES,
  };

existingSources[row.edition] =
  row.source;

editionSourcesByUnique.set(
  row.variant_id,
  existingSources,
);
}

    setUniques(
      catalogueItems.map((item) => {
        const importStatus =
          importedStatusById.get(item.id) ??
          "Unreviewed";

        return {
          id: item.id,
          familyId: item.family_id,
          name: item.name,
          baseType: item.base_type,
          itemType: item.item_type,
          variantLabel: item.variant_label,
          importStatus,
                    flags:
            flagsByUnique.get(item.id) ??
            [],
          editionAvailability:
  editionAvailabilityByUnique.get(
    item.id,
  ) ?? {
    ...DEFAULT_EDITION_AVAILABILITY,
  },
editionSources:
  editionSourcesByUnique.get(
    item.id,
  ) ?? {
    ...DEFAULT_EDITION_SOURCES,
  },
reviewed:
            reviewedByUnique.get(item.id) ??
            false,
          isLegacyOnly: item.is_legacy_only === 1,
          catalogueSource: item.source,
        };
      }),
    );
  }

  async function syncCatalogueFamiliesAndVariants(
    db: Database,
  ) {
    const catalogueRows = await db.select<
      {
        id: string;
        name: string;
        base_type: string | null;
        item_type: string;
        release_version: string | null;
        drop_enabled: number;
        drop_restricted: number;
        is_replica: number;
        has_legacy_variants: number;
        is_legacy_only: number;
        removal_version: string | null;
        source: string;
      }[]
    >(`
      SELECT
        id,
        name,
        base_type,
        item_type,
        release_version,
        drop_enabled,
        drop_restricted,
        is_replica,
        has_legacy_variants,
        is_legacy_only,
        removal_version,
        source
      FROM unique_catalogue
      WHERE source = 'poewiki'
    `);

        const existingVariantLabelRows =
      await db.select<
        {
          id: string;
          variant_label:
            string | null;
        }[]
      >(`
        SELECT
          id,
          variant_label
        FROM unique_variants
        WHERE source = 'poewiki'
      `);

    const existingVariantLabels =
      new Map(
        existingVariantLabelRows.map(
          (row) => [
            row.id,
            row.variant_label,
          ] as const,
        ),
      );

    const wikiVariantLabels =
      await fetchWikiVariantLabels(
        catalogueRows,
      );

    for (const row of catalogueRows) {
      const familyId =
        makeFamilyId(row.name);

              const rawVariantLabel =
  wikiVariantLabels.get(
    row.id,
  ) ??
  existingVariantLabels.get(
    row.id,
  ) ??
  null;

const variantLabel =
  getCuratedVariantLabel(
    row.name,
    rawVariantLabel,
  );

const isLegacyOnly =
  row.is_legacy_only === 1 ||
  isCuratedLegacyVariant(
    row.name,
    rawVariantLabel,
  );

      await db.execute(
        `
          INSERT OR REPLACE INTO unique_families (
            id,
            name,
            item_type,
            stash_slot_key,
            source
          )
          VALUES (?, ?, ?, ?, 'poewiki')
        `,
        [
          familyId,
          row.name,
          row.item_type,
          row.name.toLowerCase(),
        ],
      );

      await db.execute(
        `
          INSERT OR REPLACE INTO unique_variants (
            id,
            family_id,
            name,
            base_type,
            item_type,
            variant_label,
            release_version,
            drop_enabled,
            drop_restricted,
            is_replica,
            has_legacy_variants,
            is_legacy_only,
            removal_version,
            source
          )
                    VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'poewiki'
          )
        `,
                [
          row.id,
          familyId,
          row.name,
          row.base_type,
          row.item_type,
          variantLabel,
          row.release_version,
          row.drop_enabled,
          row.drop_restricted,
          row.is_replica,
          row.has_legacy_variants,
          isLegacyOnly ? 1 : 0,
          row.removal_version,
        ],
      );
    }

    await db.execute(`
      DELETE FROM unique_variants
      WHERE
        source = 'poewiki'
        AND id NOT IN (
          SELECT id
          FROM unique_catalogue
          WHERE source = 'poewiki'
        )
    `);

    await db.execute(`
      DELETE FROM unique_families
      WHERE id NOT IN (
        SELECT DISTINCT family_id
        FROM unique_variants
      )
    `);
  }

  async function runCatalogueCheck(
    db: Database,
    force = false,
  ) {
    if (catalogueChecking) {
      return;
    }

    try {
      const pendingLeagueDetections =
        await readPendingLeagueDetections(
          db,
        );

      const forceForNewLeague =
        pendingLeagueDetections.length > 0;

      if (
        !force &&
        !forceForNewLeague
      ) {
        const shouldCheck =
          await shouldCheckCatalogueNow(
            db,
          );

        if (!shouldCheck) {
          return;
        }
      }

      setCatalogueChecking(true);
      setAppError("");
      setCatalogueCheckMessage(
        forceForNewLeague
          ? "New league detected. Refreshing the PoE 2 Wiki catalogue..."
          : "Checking PoE 2 Wiki for catalogue updates...",
      );

      /*
       * Snapshot the tracker catalogue before the Wiki refresh.
       * This lets the normal once-a-day check identify the exact
       * entries that appeared since the previous successful scan.
       */
      const beforeCatalogue =
        await loadTrackableImportCatalogue(
          db,
        );

      const hadCatalogueBefore =
        beforeCatalogue.length > 0;

      const beforeIds =
        new Set(
          beforeCatalogue.map(
            (item) => item.id,
          ),
        );

      const result =
        await syncCatalogueFromPoeWiki(
          db,
        );

      await syncCatalogueFamiliesAndVariants(
        db,
      );
      await seedBuiltInSpecialVariants(db);

      const afterCatalogue =
        await loadTrackableImportCatalogue(
          db,
        );

      /*
       * Do not announce the entire catalogue as "new" on the
       * very first canonical sync. On established installs,
       * only IDs absent from the previous successful local
       * catalogue count as newly discovered tracker entries.
       */
      const newlyDiscoveredRows =
        hadCatalogueBefore
          ? afterCatalogue.filter(
              (item) =>
                !beforeIds.has(
                  item.id,
                ),
            )
          : [];

      const newlyDiscoveredItems:
        TrackerDiscoveryItem[] =
        newlyDiscoveredRows.map(
          (item) => ({
            id: item.id,
            name: item.name,
            itemType:
              item.item_type,
            variantLabel:
              item.variant_label,
            releaseVersion:
              item.release_version,
          }),
        );

      if (
        pendingLeagueDetections.length >
        0
      ) {
        /*
         * A catalogue entry discovered as part of a new-league
         * refresh starts as known Missing in Standard and in the
         * newly-created league collection(s). Other already-running
         * challenge profiles are left alone.
         *
         * If the Wiki had already been refreshed before the league
         * was detected, there may be zero newlyDiscoveredItems. In
         * that case we deliberately announce 0 rather than guessing
         * that an older release-line entry belongs to this league.
         */
        if (
          newlyDiscoveredItems.length >
          0
        ) {
          const profileIds =
            Array.from(
              new Set([
                STANDARD_PROFILE_ID,
                ...pendingLeagueDetections.map(
                  (entry) =>
                    entry.profileId,
                ),
              ]),
            );

          for (
            const profileId of
            profileIds
          ) {
            for (
              const item of
              newlyDiscoveredItems
            ) {
              await db.execute(
                `
                  INSERT OR REPLACE INTO profile_collection_review (
                    profile_id,
                    unique_id,
                    reviewed
                  )
                  VALUES (?, ?, 1)
                `,
                [
                  profileId,
                  item.id,
                ],
              );
            }
          }
        }

        const createdAt =
          Date.now();

        const leagueNotices:
          TrackerDiscoveryNotice[] =
          pendingLeagueDetections.map(
            (entry) => ({
              id:
                `new-league:${entry.profileId}`,
              kind:
                "new-league" as const,
              leagueName:
                entry.leagueName,
              profileId:
                entry.profileId,
              items:
                newlyDiscoveredItems,
              createdAt,
            }),
          );

        const queued =
          await appendDiscoveryNotices(
            db,
            leagueNotices,
          );

        setDiscoveryNotices(
          queued,
        );

        /*
         * Persist the finished notice before clearing the pending
         * league marker. If the app closes before this point, the
         * next launch will retry the catalogue refresh instead of
         * losing the league announcement.
         */
        await clearPendingLeagueDetections(
          db,
        );
      } else if (
        newlyDiscoveredItems.length >
        0
      ) {
        /*
         * A unique discovered during an already-running league is
         * intentionally left without profile_collection_review rows.
         * That means it appears as Unreviewed in every collection
         * until the player chooses a status.
         */
        const queued =
          await appendDiscoveryNotices(
            db,
            [
              {
                id:
                  `new-unique:${result.revision}`,
                kind:
                  "new-unique",
                items:
                  newlyDiscoveredItems,
                createdAt:
                  Date.now(),
              },
            ],
          );

        setDiscoveryNotices(
          queued,
        );
      }

      const reconciliation =
        await reconcileImportedCollection(
          db,
        );

      setImportMatchSummary(
        reconciliation,
      );

      await loadCollectionData(
        db,
        activeProfileId,
      );

      const checkedTime =
        new Date(
          result.checkedAt,
        ).toLocaleString();

      setCatalogueCheckMessage(
        `Last checked ${checkedTime} \u2022 ${result.remoteItems.toLocaleString()} catalogue entries received.`,
      );

      /*
       * New entries get their own player-facing discovery popup.
       * Keep the older generic catalogue-update popup for metadata
       * changes that do not add a new tracker entry.
       */
      if (
        result.hasVisibleChanges &&
        pendingLeagueDetections.length ===
          0 &&
        newlyDiscoveredItems.length ===
          0
      ) {
        setCatalogueUpdate({
          revision:
            result.revision,
          label:
            result.label,
          newFamilies:
            result.newFamilies,
          newVariants:
            result.newVariants,
          dropDisabled:
            result.dropDisabled,
          updatedEntries:
            result.updatedEntries,
        });
        setCatalogueUpdateOpen(
          true,
        );
      }
    } catch (error) {
      console.error(
        "Could not update catalogue from PoE 2 Wiki:",
        error,
      );

      const message =
        error instanceof Error
          ? error.message
          : String(error);

      setCatalogueCheckMessage(
        `Catalogue check failed: ${message}`,
      );

    } finally {
      setCatalogueChecking(false);
    }
  }


  const activeDiscoveryNotice =
    devDiscoveryNotice ??
    discoveryNotices[0] ??
    null;

  async function dismissDiscoveryNotice() {
    if (devDiscoveryNotice) {
      setDevDiscoveryNotice(null);
      setDiscoveryDetailsExpanded(
        false,
      );
      return;
    }

    if (
      !database ||
      !activeDiscoveryNotice
    ) {
      return;
    }

    const remaining =
      discoveryNotices.slice(1);

    try {
      await saveDiscoveryNotices(
        database,
        remaining,
      );

      setDiscoveryNotices(
        remaining,
      );
      setDiscoveryDetailsExpanded(
        false,
      );
    } catch (error) {
      console.error(
        "Could not dismiss tracker discovery notice:",
        error,
      );
    }
  }

  function showDevNewLeagueNotice() {
    setSettingsOpen(false);
    setDiscoveryDetailsExpanded(false);

    setDevDiscoveryNotice({
      id: "dev:new-league-popup",
      kind: "new-league",
      leagueName: "Test League",
      profileId: "dev:test-league",
      createdAt: Date.now(),
      items: [
        {
          id: "dev:league-unique-1",
          name: "Example Unique One",
          itemType: "Ring",
          variantLabel: null,
          releaseVersion: "0.6",
        },
        {
          id: "dev:league-unique-2",
          name: "Example Unique Two",
          itemType: "Helmet",
          variantLabel: null,
          releaseVersion: "0.6",
        },
        {
          id: "dev:league-unique-3",
          name: "Example Unique Three",
          itemType: "Sword",
          variantLabel: null,
          releaseVersion: "0.6",
        },
        {
          id: "dev:league-unique-4",
          name: "Example Unique Four",
          itemType: "Amulet",
          variantLabel: "Alternate Variant",
          releaseVersion: "0.6",
        },
        {
          id: "dev:league-unique-5",
          name: "Example Unique Five",
          itemType: "Body Armour",
          variantLabel: null,
          releaseVersion: "0.6",
        },
      ],
    });
  }

  function showDevMidLeagueUniqueNotice() {
    setSettingsOpen(false);
    setDiscoveryDetailsExpanded(false);

    setDevDiscoveryNotice({
      id: "dev:mid-league-unique-popup",
      kind: "new-unique",
      createdAt: Date.now(),
      items: [
        {
          id: "dev:mid-league-unique-1",
          name: "Example Newly Discovered Unique",
          itemType: "Quarterstaff",
          variantLabel: null,
          releaseVersion: "0.5",
        },
      ],
    });
  }

  async function acknowledgeCatalogueUpdate(reviewNewItems: boolean) {
    if (!database || !catalogueUpdate) {
      setCatalogueUpdateOpen(false);
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('acknowledged_catalogue_revision', ?)
        `,
        [catalogueUpdate.revision],
      );

      setCatalogueUpdateOpen(false);

      if (reviewNewItems) {
        setSearchTerm("");
        setTypeFilter("All");
        setStatusFilter("unreviewed");
      }
    } catch (error) {
      console.error("Could not acknowledge catalogue update:", error);
    }
  }

  async function changeStatusColor(status: StatusKey, color: string) {
    const newColors = {
      ...statusColors,
      [status]: color,
    };

    setStatusColors(newColors);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('status_colors', ?)
        `,
        [JSON.stringify(newColors)],
      );
    } catch (error) {
      console.error("Could not save status colors:", error);
    }
  }

  async function resetStatusColors() {
    setStatusColors(DEFAULT_STATUS_COLORS);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('status_colors', ?)
        `,
        [JSON.stringify(DEFAULT_STATUS_COLORS)],
      );
    } catch (error) {
      console.error("Could not reset status colors:", error);
    }
  }

  async function changeCollectionRule(flag: TrackingFlag, enabled: boolean) {
    const newRules = {
      ...collectionRules,
      [flag]: enabled,
    };

    setCollectionRules(newRules);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('collection_rules', ?)
        `,
        [JSON.stringify(newRules)],
      );
    } catch (error) {
      console.error("Could not save collection rules:", error);
    }
  }

  async function changeExtraTracking(
    variant: ExtraVariant,
    enabled: boolean,
  ) {
    const newTracking = {
      ...extraTracking,
      [variant]: enabled,
    };

    setExtraTracking(newTracking);

    if (!enabled && statusFilter === variant) {
      setStatusFilter("all");
    }

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('extra_tracking', ?)
        `,
        [JSON.stringify(newTracking)],
      );
    } catch (error) {
      console.error("Could not save extra tracking settings:", error);
    }
  }

  async function changeShowLegacyVariants(
    enabled: boolean,
  ) {
    setShowLegacyVariants(enabled);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('show_legacy_variants', ?)
        `,
        [enabled ? "true" : "false"],
      );
    } catch (error) {
      console.error(
        "Could not save legacy display preference:",
        error,
      );
    }
  }

  async function activateLookupHotkey(
    newHotkey: string,
  ) {
    const currentHotkey =
      registeredLookupHotkeyRef.current;

    if (currentHotkey === newHotkey) {
      setHotkeyReady(true);
      setHotkeyMessage(
        `${formatHotkeyForDisplay(newHotkey)} is ready.`,
      );
      return;
    }

    try {
      /*
       * Hot reload can leave a stale registration owned by this app.
       * unregister() only affects PoE 2 Unique Tracker's own registration.
       */
      try {
        await unregister(newHotkey);
      } catch {
        // Fine if PoE 2 Unique Tracker did not have it registered.
      }

      /*
       * Register the replacement BEFORE releasing the currently-working
       * shortcut. If Windows rejects it because another app owns it, the
       * player's existing hotkey remains active.
       */
      await register(
        newHotkey,
        (event) => {
          if (event.state === "Released") {
            void hotkeyLookupRef.current();
          }
        },
      );

      if (
        currentHotkey &&
        currentHotkey !== newHotkey
      ) {
        try {
          await unregister(currentHotkey);
        } catch (error) {
          console.warn(
            "Could not unregister the previous lookup hotkey:",
            error,
          );
        }
      }

      registeredLookupHotkeyRef.current =
        newHotkey;

      setHotkeyReady(true);
      setHotkeyMessage(
        `${formatHotkeyForDisplay(newHotkey)} is ready.`,
      );
    } catch (error) {
      setHotkeyReady(
        registeredLookupHotkeyRef.current !== null,
      );

      const previous =
        registeredLookupHotkeyRef.current;

      setHotkeyMessage(
        previous
          ? `${formatHotkeyForDisplay(previous)} is still active. The new shortcut could not be registered.`
          : error instanceof Error
            ? `Hotkey unavailable: ${error.message}`
            : `Hotkey unavailable: ${String(error)}`,
      );

      throw error;
    }
  }

  async function changeLookupHotkey(
    newHotkey: string,
  ) {
    if (!database) {
      return;
    }

    if (newHotkey === lookupHotkey) {
      setHotkeyRecording(false);
      setHotkeySettingMessage("");
      return;
    }

    if (
      import.meta.env.DEV &&
      newHotkey ===
        DEV_LEAGUE_ROLLOVER_HOTKEY
    ) {
      setHotkeySettingMessage(
        "That shortcut is reserved for the development league-rollover test.",
      );
      return;
    }

    const previousHotkey =
      lookupHotkey;

    try {
      setAppError("");
      setHotkeySettingMessage(
        "Checking shortcut...",
      );

      /*
       * Prove the shortcut actually works before persisting it.
       * This prevents a conflicting shortcut from replacing a known-good
       * saved hotkey.
       */
      await activateLookupHotkey(
        newHotkey,
      );

      try {
        await database.execute(
          `
            INSERT OR REPLACE INTO app_meta (
              key,
              value
            )
            VALUES (
              'poe_lookup_hotkey',
              ?
            )
          `,
          [newHotkey],
        );
      } catch (error) {
        /*
         * The new shortcut was activated but could not be saved.
         * Restore the previous one so runtime state and saved state agree.
         */
        await activateLookupHotkey(
          previousHotkey,
        ).catch(
          (restoreError) => {
            console.error(
              "Could not restore the previous lookup hotkey:",
              restoreError,
            );
          },
        );

        throw error;
      }

      setLookupHotkey(newHotkey);
      setHotkeyRecording(false);
      setHotkeySettingMessage(
        `Hotkey changed to ${formatHotkeyForDisplay(newHotkey)}.`,
      );
    } catch (error) {
      console.error(
        "Could not change lookup hotkey:",
        error,
      );

      setHotkeySettingMessage(
        error instanceof Error
          ? `Could not use that shortcut: ${error.message}`
          : `Could not use that shortcut: ${String(error)}`,
      );
    }
  }

function captureLookupHotkey(
  event:
    ReactKeyboardEvent<HTMLButtonElement>,
) {
  event.preventDefault();
  event.stopPropagation();

  if (event.key === "Escape") {
    setHotkeyRecording(false);
    setHotkeySettingMessage("");
    return;
  }

  if (
    [
      "Control",
      "Shift",
      "Alt",
      "Meta",
    ].includes(event.key)
  ) {
    return;
  }

  const key =
    normalizeCapturedHotkeyKey(
      event.key,
    );

  if (!key) {
    setHotkeySettingMessage(
      "Use a letter, number, or F1\u2013F12 as the main key.",
    );
    return;
  }

  const parts: string[] = [];

  if (event.ctrlKey) {
    parts.push(
      "CommandOrControl",
    );
  }

  if (event.altKey) {
    parts.push("Alt");
  }

  if (event.shiftKey) {
    parts.push("Shift");
  }

  const isFunctionKey =
    /^F([1-9]|1[0-2])$/.test(key);

  if (
    parts.length === 0 &&
    !isFunctionKey
  ) {
    setHotkeySettingMessage(
      "Use Ctrl, Alt, or Shift with letters and numbers.",
    );
    return;
  }

  parts.push(key);

  void changeLookupHotkey(
    parts.join("+"),
  );
}

  useEffect(() => {
  if (!database) {
    return;
  }

  void ensureOverlayWindow().catch(
    (error) => {
      console.error(
        "Could not prepare overlay window:",
        error,
      );
    },
  );
}, [database]);

useEffect(() => {
  if (!database) {
    return;
  }

  void activateLookupHotkey(
    lookupHotkey,
  ).catch((error) => {
    console.error(
      "Could not register global shortcut:",
      error,
    );
  });
}, [database, lookupHotkey]);

useEffect(() => {
  if (
    !database ||
    !import.meta.env.DEV
  ) {
    return;
  }

  let cancelled = false;

  void (async () => {
    try {
      try {
        await unregister(
          DEV_LEAGUE_ROLLOVER_HOTKEY,
        );
      } catch {
        // Fine if it was not registered.
      }

      await register(
        DEV_LEAGUE_ROLLOVER_HOTKEY,
        (event) => {
          if (
            !cancelled &&
            event.state === "Released"
          ) {
            void previewLeagueRollover();
          }
        },
      );
    } catch (error) {
      console.error(
        "Could not register development rollover shortcut:",
        error,
      );
    }
  })();

  return () => {
    cancelled = true;

    void unregister(
      DEV_LEAGUE_ROLLOVER_HOTKEY,
    ).catch(() => {});
  };
}, [database]);

useEffect(() => {
  return () => {
    const currentHotkey =
      registeredLookupHotkeyRef.current;

    if (currentHotkey) {
      void unregister(
        currentHotkey,
      ).catch(() => {});
    }
  };
}, []);

  useEffect(() => {
  let unlistenSearch:
    | (() => void)
    | undefined;

  void listen<string>(
    "poe-overlay-search-item",
    (event) => {
      setSearchTerm(event.payload);
      setStatusFilter("all");
      setTypeFilter("All");
    },
  ).then((unlisten) => {
    unlistenSearch = unlisten;
  });

  return () => {
    unlistenSearch?.();
  };
}, []);

  useEffect(() => {
    if (!database) {
      return;
    }

    let unlistenAction:
      | (() => void)
      | undefined;

    void listen<OverlayAction>(
      "poe-overlay-action",
      (event) => {
        const action = event.payload;

        if (action.kind === "missing") {
          void markUniqueMissing(
            action.uniqueId,
            action.profileId,
          );
          return;
        }

        void applyOverlayFlag(
          action.profileId,
          action.uniqueId,
          action.flag,
          action.enabled,
        );
      },
      {
        target: { kind: "Any" },
      },
    ).then((unlisten) => {
      unlistenAction = unlisten;
    });

    return () => {
      unlistenAction?.();
    };
  }, [database, activeProfileId]);

  useEffect(() => {
    async function initializeDatabase() {
      try {
        const db = await Database.load(
  "sqlite:poe2-collector.db",
);

await db.execute(
  "PRAGMA journal_mode = WAL",
);

await db.execute(
  "PRAGMA synchronous = NORMAL",
);

await db.execute(
  "PRAGMA busy_timeout = 15000",
);

await ensureEditionAvailabilitySchema(
  db,
);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS imported_collection (
            id TEXT PRIMARY KEY NOT NULL,
            name TEXT NOT NULL,
            item_type TEXT NOT NULL,
            status TEXT NOT NULL
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS app_meta (
            key TEXT PRIMARY KEY NOT NULL,
            value TEXT NOT NULL
          )
        `);

        const savedDiscoveryNotices =
          await readDiscoveryNotices(
            db,
          );

        setDiscoveryNotices(
          savedDiscoveryNotices,
        );

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_tracking (
            unique_id TEXT NOT NULL,
            flag TEXT NOT NULL,
            PRIMARY KEY (unique_id, flag)
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_catalogue (
            id TEXT PRIMARY KEY NOT NULL,
            name TEXT NOT NULL,
            base_type TEXT,
            item_type TEXT NOT NULL,
            release_version TEXT,
            drop_enabled INTEGER NOT NULL DEFAULT 1,
            drop_restricted INTEGER NOT NULL DEFAULT 0,
            is_replica INTEGER NOT NULL DEFAULT 0,
            has_legacy_variants INTEGER NOT NULL DEFAULT 0,
            is_legacy_only INTEGER NOT NULL DEFAULT 0,
            removal_version TEXT,
            source TEXT NOT NULL DEFAULT 'imported'
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_families (
            id TEXT PRIMARY KEY NOT NULL,
            name TEXT NOT NULL,
            item_type TEXT NOT NULL,
            stash_slot_key TEXT,
            source TEXT NOT NULL DEFAULT 'imported'
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_variants (
            id TEXT PRIMARY KEY NOT NULL,
            family_id TEXT NOT NULL,
            name TEXT NOT NULL,
            base_type TEXT,
            item_type TEXT NOT NULL,
            variant_label TEXT,
            release_version TEXT,
            drop_enabled INTEGER NOT NULL DEFAULT 1,
            drop_restricted INTEGER NOT NULL DEFAULT 0,
            is_replica INTEGER NOT NULL DEFAULT 0,
            has_legacy_variants INTEGER NOT NULL DEFAULT 0,
            is_legacy_only INTEGER NOT NULL DEFAULT 0,
            removal_version TEXT,
            source TEXT NOT NULL DEFAULT 'imported'
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS collection_review (
            unique_id TEXT PRIMARY KEY NOT NULL,
            reviewed INTEGER NOT NULL DEFAULT 0
          )
        `);

        await db.execute(`
          INSERT OR IGNORE INTO collection_review (
            unique_id,
            reviewed
          )
          SELECT id, 1
          FROM imported_collection
        `);

        const addedSpecialVariants =
  await seedBuiltInSpecialVariants(db);

const canonicalReady =
  await isCanonicalCatalogueReady(db);

if (canonicalReady) {
  const reconciliation =
    await loadImportReconciliationSummary(db);

  setImportMatchSummary(reconciliation);
}

        const migrationState = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'tracking_flags_migrated_v1'
        `);

        if (migrationState.length === 0) {
          await db.execute(`
            INSERT OR IGNORE INTO unique_tracking (unique_id, flag)
            SELECT id, 'owned'
            FROM imported_collection
            WHERE status = 'Owned'
          `);

          await db.execute(`
            INSERT OR IGNORE INTO unique_tracking (unique_id, flag)
            SELECT id, 'wearing'
            FROM imported_collection
            WHERE status = 'Wearing'
          `);

          await db.execute(`
            INSERT OR IGNORE INTO unique_tracking (unique_id, flag)
            SELECT id, 'foil'
            FROM imported_collection
            WHERE status IN ('Owned Foil', 'Owned Rainbow')
          `);

          await db.execute(`
            INSERT OR REPLACE INTO app_meta (key, value)
            VALUES ('tracking_flags_migrated_v1', 'yes')
          `);
        }

        const profileState =
  await initializeCollectionProfiles(db);

setCollectionProfiles(
  profileState.profiles,
);
setActiveProfileId(
  profileState.activeProfileId,
);

        const savedSource = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'source_file'
        `);

        const savedImportMode =
          await db.select<
            { value: string }[]
          >(`
            SELECT value
            FROM app_meta
            WHERE key = 'source_import_mode'
          `);

        const savedSortMode = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'sort_mode'
        `);

        const savedStatusColors = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'status_colors'
        `);

        const savedCollectionRules = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'collection_rules'
        `);

        const savedExtraTracking = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'extra_tracking'
        `);

        const savedShowLegacyVariants = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'show_legacy_variants'
        `);

        const savedLookupHotkey =
  await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'poe_lookup_hotkey'
  `);

  const hotkeyMigrationState =
  await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'poe2_hotkey_default_migrated_v1'
  `);

        if (savedStatusColors.length > 0) {
          try {
            const parsedColors = JSON.parse(savedStatusColors[0].value);

            setStatusColors({
              ...DEFAULT_STATUS_COLORS,
              ...parsedColors,
            });
          } catch (error) {
            console.error("Could not load saved status colors:", error);
          }
        }

        if (savedCollectionRules.length > 0) {
          try {
            const parsedRules = JSON.parse(savedCollectionRules[0].value);

            setCollectionRules({
              ...DEFAULT_COLLECTION_RULES,
              ...parsedRules,
            });
          } catch (error) {
            console.error("Could not load collection rules:", error);
          }
        }

        if (savedExtraTracking.length > 0) {
          try {
            const parsedTracking = JSON.parse(savedExtraTracking[0].value);

            setExtraTracking({
              ...DEFAULT_EXTRA_TRACKING,
              ...parsedTracking,
            });
          } catch (error) {
            console.error("Could not load extra tracking settings:", error);
          }
        }

        if (savedShowLegacyVariants.length > 0) {
          setShowLegacyVariants(
            savedShowLegacyVariants[0].value !== "false",
          );
        }

        let resolvedLookupHotkey =
  savedLookupHotkey[0]?.value.trim() ||
  DEFAULT_POE_LOOKUP_HOTKEY;

if (
  hotkeyMigrationState.length === 0
) {
  if (
    resolvedLookupHotkey ===
    OLD_DEFAULT_POE_LOOKUP_HOTKEY
  ) {
    resolvedLookupHotkey =
      DEFAULT_POE_LOOKUP_HOTKEY;

    await db.execute(
      `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'poe_lookup_hotkey',
          ?
        )
      `,
      [resolvedLookupHotkey],
    );
  }

  await db.execute(`
    INSERT OR REPLACE INTO app_meta (
      key,
      value
    )
    VALUES (
      'poe2_hotkey_default_migrated_v1',
      'yes'
    )
  `);
}

setLookupHotkey(
  resolvedLookupHotkey,
);

        await loadCollectionData(
          db,
          profileState.activeProfileId,
        );

        if (savedSource.length > 0) {
          setSourceFile(savedSource[0].value);
        }

        const savedImportModeValue =
          savedImportMode[0]?.value;

        if (
          savedImportModeValue ===
            "status-list" ||
          savedImportModeValue ===
            "missing-only" ||
          savedImportModeValue ===
            "color-coded-list"
        ) {
          setSourceImportMode(
            savedImportModeValue,
          );
        }
        else if (
          savedImportModeValue ===
          "ole-color-list"
        ) {
          // Compatibility with pre-release color-coded import tests.
          setSourceImportMode(
            "color-coded-list",
          );
        }

        if (
          savedSortMode.length > 0 &&
          (savedSortMode[0].value === "alphabetical" ||
            savedSortMode[0].value === "type")
        ) {
          setSortMode(savedSortMode[0].value as SortMode);
        }

        setDatabase(db);
        setDatabaseReady(true);

        if (addedSpecialVariants > 0) {
          setCatalogueUpdate({
            revision: "special-variants-v1",
            label: "Special Variant Catalogue Update",
            newFamilies: 0,
            newVariants: addedSpecialVariants,
            dropDisabled: 0,
            updatedEntries: 0,
          });
          setCatalogueUpdateOpen(true);
        }

        /*
 * League detection and catalogue syncing can both write
 * to SQLite. Run them one after the other so they do not
 * compete for the database during startup.
 *
 * If a new league was created, initializeLiveLeagues()
 * leaves a persisted pending marker. runCatalogueCheck()
 * sees that marker and forces a fresh Wiki scan even when
 * the normal once-a-day catalogue check is not due yet.
 */
void (async () => {
  const liveProfiles =
    await initializeLiveLeagues(
      db,
      profileState.profiles,
    );

  setCollectionProfiles(
    liveProfiles,
  );

  await runCatalogueCheck(
    db,
    false,
  );
})();
      } catch (error) {
        console.error(error);

        setAppError(
          error instanceof Error ? error.message : String(error),
        );
      }
    }

    initializeDatabase();
  }, []);

  const visibleUniques = useMemo(
    () =>
      showLegacyVariants
        ? uniques
        : uniques.filter(
            (unique) => !unique.isLegacyOnly,
          ),
    [uniques, showLegacyVariants],
  );

  const trackingCounts = useMemo(() => {
    return [
      {
        label: "Owned",
        count: visibleUniques.filter(
          (unique) =>
            isUniqueCollected(unique),
        ).length,
      },
      {
        label: "Missing",
        count: visibleUniques.filter((unique) =>
          isUniqueMissing(unique),
        ).length,
      },
      {
        label: "Unreviewed",
        count: visibleUniques.filter((unique) => !unique.reviewed).length,
      },
      {
        label: "Wearing",
        count: visibleUniques.filter((unique) => unique.flags.includes("wearing"))
          .length,
      },
      ...(extraTracking.foil
        ? [
            {
              label: "Foil",
              count: visibleUniques.filter((unique) =>
                unique.flags.includes("foil"),
              ).length,
            },
          ]
        : []),
      ...(extraTracking.foulborn
        ? [
            {
              label: "Foulborn",
              count: visibleUniques.filter((unique) =>
                unique.flags.includes("foulborn"),
              ).length,
            },
          ]
        : []),
      ...(extraTracking.vestigial
        ? [
            {
              label: "Vestigial",
              count: visibleUniques.filter((unique) =>
                unique.flags.includes("vestigial"),
              ).length,
            },
          ]
        : []),
    ];
  }, [visibleUniques, collectionRules, extraTracking]);

  const availableTypes = useMemo(() => {
    const types = new Set<string>(
      visibleUniques.map((unique) => unique.itemType),
    );

    return Array.from(types).sort(
      comparePoe2UniqueTabTypes,
    );
  }, [visibleUniques]);

  const displayedUniques = useMemo(() => {
    const normalizedSearch = searchTerm.trim().toLowerCase();

    const filtered = visibleUniques.filter((unique) => {
      const searchableText = [
        unique.name,
        unique.baseType ?? "",
        unique.variantLabel ?? "",
      ]
        .join(" ")
        .toLowerCase();

      const matchesSearch =
        normalizedSearch === "" ||
        searchableText.includes(normalizedSearch);

      let matchesStatus = true;

      if (statusFilter === "unreviewed") {
        matchesStatus = !unique.reviewed;
      } else if (statusFilter === "missing") {
        matchesStatus = isUniqueMissing(unique);
      } else if (statusFilter === "owned") {
        matchesStatus = isUniqueCollected(unique);
      } else if (statusFilter !== "all") {
        matchesStatus = unique.flags.includes(statusFilter);
      }

      const matchesType =
        typeFilter === "All" || unique.itemType === typeFilter;

      return matchesSearch && matchesStatus && matchesType;
    });

    return [...filtered].sort((a, b) => {
      if (sortMode === "type") {
        const typeComparison =
          comparePoe2UniqueTabTypes(
            a.itemType,
            b.itemType,
          );

        if (typeComparison !== 0) {
          return typeComparison;
        }
      }

      const nameComparison = a.name.localeCompare(b.name);

      if (nameComparison !== 0) {
        return nameComparison;
      }

      return (a.variantLabel ?? "").localeCompare(
        b.variantLabel ?? "",
      );
    });
  }, [
    visibleUniques,
    searchTerm,
    statusFilter,
    typeFilter,
    sortMode,
    collectionRules,
  ]);

  const hasUncertainEditionData =
  useMemo(
    () =>
      displayedUniques.some(
        (unique) =>
          EXTRA_VARIANTS.some(
            (edition) =>
              shouldShowEdition(
                edition,
                unique.editionAvailability,
                unique.flags,
                extraTracking,
              ) &&
              isEditionUncertain(
                edition,
                unique.editionAvailability,
                unique.editionSources,
                unique.flags,
              ),
          ),
      ),
    [
      displayedUniques,
      extraTracking,
    ],
  );

  const groupedUniques = useMemo(() => {
    const groups = new Map<string, UniqueEntry[]>();

    for (const unique of displayedUniques) {
      const existing = groups.get(unique.itemType) ?? [];
      existing.push(unique);
      groups.set(unique.itemType, existing);
    }

    return Array.from(groups.entries()).map(([itemType, items]) => ({
      itemType,
      items,
    }));
  }, [displayedUniques]);

  async function ensureOverlayWindow() {
  const existing =
    await WebviewWindow.getByLabel(
      OVERLAY_LABEL,
    );

  if (existing) {
    return existing;
  }

  const overlay = new WebviewWindow(
    OVERLAY_LABEL,
    {
      url: "index.html?overlay=1",
      title: "PoE 2 Unique Tracker Lookup",
      width: OVERLAY_WIDTH,
      height: OVERLAY_HEIGHT,
      decorations: false,
      resizable: false,
      maximizable: false,
      minimizable: false,
      alwaysOnTop: true,
      skipTaskbar: true,
      focus: false,
      visible: false,
      shadow: true,
    },
  );

  await wait(100);

  return overlay;
}

async function showOverlayLoading(
  capturedCursor: {
    x: number;
    y: number;
  } | null,
) {
  const overlay =
    await ensureOverlayWindow();

  let monitor = null;

  try {
    if (capturedCursor) {
      monitor = await monitorFromPoint(
        capturedCursor.x,
        capturedCursor.y,
      );
    }
  } catch {
    monitor = null;
  }

  if (!monitor) {
    monitor = await primaryMonitor();
  }

  if (monitor) {
    const workPosition =
      monitor.workArea.position.toLogical(
        monitor.scaleFactor,
      );

    const workSize =
      monitor.workArea.size.toLogical(
        monitor.scaleFactor,
      );

    const cursorLogical = capturedCursor
      ? {
          x:
            capturedCursor.x /
            monitor.scaleFactor,
          y:
            capturedCursor.y /
            monitor.scaleFactor,
        }
      : {
          x:
            workPosition.x +
            workSize.width / 2,
          y:
            workPosition.y +
            workSize.height / 2,
        };

    const gap = 18;

    let x =
      cursorLogical.x + gap;

    let y =
      cursorLogical.y + gap;

    const rightEdge =
      workPosition.x +
      workSize.width;

    const bottomEdge =
      workPosition.y +
      workSize.height;

    if (
      x + OVERLAY_WIDTH >
      rightEdge - 8
    ) {
      x =
        cursorLogical.x -
        OVERLAY_WIDTH -
        gap;
    }

    if (
      y + OVERLAY_HEIGHT >
      bottomEdge - 8
    ) {
      y =
        cursorLogical.y -
        OVERLAY_HEIGHT -
        gap;
    }

    x = Math.max(
      workPosition.x + 8,
      Math.min(
        x,
        rightEdge -
          OVERLAY_WIDTH -
          8,
      ),
    );

    y = Math.max(
      workPosition.y + 8,
      Math.min(
        y,
        bottomEdge -
          OVERLAY_HEIGHT -
          8,
      ),
    );

    await overlay.setPosition(
      new LogicalPosition(x, y),
    );
  }

  await emit("poe-overlay-loading");

  // One frame so "Checking..." paints before
  // the hidden window becomes visible.
  await wait(16);

  await overlay.show();
}

  async function applyOverlayFlag(
    profileId: string,
    uniqueId: string,
    flag: TrackingFlag,
    enabled: boolean,
  ) {
    if (!database) {
      return;
    }

    const enabledFlags: TrackingFlag[] =
      enabled && flag === "foil"
        ? ["owned", "foil"]
        : [flag];

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        [profileId, uniqueId],
      );

      if (enabled) {
        for (const enabledFlag of enabledFlags) {
          await database.execute(
            `
              INSERT OR REPLACE INTO profile_unique_tracking (
                profile_id,
                unique_id,
                flag
              )
              VALUES (?, ?, ?)
            `,
            [
              profileId,
              uniqueId,
              enabledFlag,
            ],
          );
        }
      } else {
        await database.execute(
          `
            DELETE FROM profile_unique_tracking
            WHERE
              profile_id = ?
              AND unique_id = ?
              AND flag = ?
          `,
          [
            profileId,
            uniqueId,
            flag,
          ],
        );
      }

      if (profileId === activeProfileId) {
        setUniques((current) =>
          current.map((item) =>
            item.id === uniqueId
              ? {
                  ...item,
                  reviewed: true,
                  flags: enabled
                    ? Array.from(
                        new Set([
                          ...item.flags,
                          ...enabledFlags,
                        ]),
                      )
                    : item.flags.filter(
                        (existingFlag) =>
                          existingFlag !== flag,
                      ),
                }
              : item,
          ),
        );
      }
    } catch (error) {
      console.error(
        "Could not apply overlay status:",
        error,
      );
    }
  }

  async function openItemOverlay(
  result: Extract<
    ItemIdentificationResult,
    { status: "matched" }
  >,
) {
    if (!database) {
      return;
    }

    const localEntry = uniques.find(
      (unique) => unique.id === result.uniqueId,
    );

    if (!localEntry) {
      setHotkeyMessage(
        `Matched ${result.name}, but its collection row could not be loaded.`,
      );
      return;
    }

    const profileStateRows =
  await database.select<
    {
      id: string;
      name: string;
      flag: string | null;
    }[]
  >(
    `
      SELECT
        profiles.id,
        profiles.name,
        tracking.flag
      FROM collection_profiles profiles
      LEFT JOIN profile_unique_tracking tracking
        ON
          tracking.profile_id = profiles.id
          AND tracking.unique_id = ?
      WHERE profiles.is_archived = 0
      ORDER BY
        profiles.sort_order ASC,
        profiles.name ASC
    `,
    [localEntry.id],
  );

const profileStatesById =
  new Map<
    string,
    OverlayProfileState
  >();

for (const row of profileStateRows) {
  let profileState =
    profileStatesById.get(row.id);

  if (!profileState) {
    profileState = {
      id: row.id,
      name: row.name,
      flags: [],
    };

    profileStatesById.set(
      row.id,
      profileState,
    );
  }

  if (
    row.flag &&
    isTrackingFlag(row.flag)
  ) {
    profileState.flags.push(
      row.flag,
    );
  }
}

const profileStates =
  Array.from(
    profileStatesById.values(),
  );

    const payload: OverlayPayload = {
      trackingProfileId: activeProfileId,
      trackingProfileName:
        activeProfile?.name ?? "Standard",
      uniqueId: localEntry.id,
      name: result.name,
      baseType: result.baseType,
      itemType: result.itemType,
      variantLabel: result.variantLabel,
      edition: result.edition,
      profileStates,
      trackingFlags: [...localEntry.flags],
      trackingReviewed: localEntry.reviewed,
      editionAvailability:
  localEntry.editionAvailability,
editionSources:
  localEntry.editionSources,
isLegacyOnly: localEntry.isLegacyOnly,
      statusColors,
      collectionRules,
      extraTracking,
    };

    const preparedOverlay =
  await ensureOverlayWindow();

await emit<OverlayPayload>(
  "poe-overlay-result",
  payload,
);

await preparedOverlay.show();
await preparedOverlay.setFocus();

return;

  }

  async function identifyHoveredItemFromHotkey() {
    if (!database || hotkeyBusy) {
      return;
    }

    setHotkeyBusy(true);
    setHotkeyMessage("Copying hovered Path of Exile 2 item...");

    let previousClipboard = "";
    let capturedCursor: {
      x: number;
      y: number;
    } | null = null;

    try {
      try {
        const cursor = await cursorPosition();

        capturedCursor = {
          x: cursor.x,
          y: cursor.y,
        };
      } catch {
  capturedCursor = null;
}

try {
  previousClipboard = await readText();
      } catch {
        previousClipboard = "";
      }

      const sentinel =
        `__POE_COLLECTOR_WAITING_${Date.now()}__`;

      await writeText(sentinel);

      // The shortcut callback fires on release. Give Windows a brief moment
      // to finish releasing the physical Ctrl/Shift keys before we inject
      // the ordinary Ctrl+C that Path of Exile 2 understands.
      await wait(35);

      await invoke("send_ctrl_c");

      let clipboardText = sentinel;

      for (let attempt = 0; attempt < 30; attempt += 1) {
        await wait(10);
        clipboardText = await readText();

        if (clipboardText !== sentinel) {
          break;
        }
      }

      await showOverlayLoading(
        capturedCursor,
      );

      if (
        !clipboardText.trim() ||
        clipboardText === sentinel
      ) {
        await writeText(previousClipboard);

        setHotkeyMessage(
          "No item text was copied. Make sure the mouse is hovering a Path of Exile 2 item.",
        );

        await emit<string>(
          "poe-overlay-message",
          "No Path of Exile 2 item was copied.",
        );

        return;
      }

      setParserTestText(clipboardText);

      const result = await identifyUniqueFromClipboard(
        database,
        clipboardText,
      );

      setParserTestResult(result);

      if (result.status === "matched") {
        const editionText =
          result.edition === "normal"
            ? ""
            : ` \u2022 ${result.edition}`;

        setHotkeyMessage(
          `Matched ${result.name}${editionText}.`,
        );

        await openItemOverlay(
          result,
        );
      } else {
        setHotkeyMessage(
          "The hotkey copied an item, but PoE 2 Unique Tracker could not match it cleanly.",
        );

        await emit<string>(
          "poe-overlay-message",
         "Item copied, but no clean catalogue match was found.",
        );
        const messageOverlay =
          await WebviewWindow.getByLabel(
            OVERLAY_LABEL,
          );

        if (messageOverlay) {
          await messageOverlay.setFocus();
        }
      }
    } catch (error) {
      console.error("Global PoE lookup failed:", error);

      if (previousClipboard) {
        try {
          await writeText(previousClipboard);
        } catch {
          // Restoring the old clipboard is best-effort only.
        }
      }

      const lookupErrorMessage =
        error instanceof Error
          ? `Lookup failed: ${error.message}`
          : `Lookup failed: ${String(error)}`;

      setHotkeyMessage(lookupErrorMessage);

      try {
        await emit<string>(
          "poe-overlay-message",
          lookupErrorMessage,
        );
      } catch {
        // The overlay may not have been created yet.
      }
    } finally {
      setHotkeyBusy(false);
    }
  }

  hotkeyLookupRef.current =
    identifyHoveredItemFromHotkey;

  async function readClipboardAndIdentify() {
    if (!database) {
      return;
    }

    try {
      setParserTesting(true);
      setParserTestResult(null);

      const clipboardText = await readText();

      if (!clipboardText.trim()) {
        setParserTestText("");
        setParserTestResult({
          status: "invalid",
          parsed: {
            itemClass: null,
            rarity: null,
            name: null,
            baseType: null,
            canonicalName: null,
            canonicalBaseType: null,
            edition: "normal",
            rawText: "",
          },
          message: "The clipboard does not contain any text.",
        });
        return;
      }

      setParserTestText(clipboardText);

      const result = await identifyUniqueFromClipboard(
        database,
        clipboardText,
      );

      setParserTestResult(result);
    } catch (error) {
      console.error("Could not read and identify clipboard:", error);

      setParserTestResult({
        status: "invalid",
        parsed: {
          itemClass: null,
          rarity: null,
          name: null,
          baseType: null,
          canonicalName: null,
          canonicalBaseType: null,
          edition: "normal",
          rawText: "",
        },
        message:
          error instanceof Error ? error.message : String(error),
      });
    } finally {
      setParserTesting(false);
    }
  }

  async function runParserTest() {
    if (!database || !parserTestText.trim()) {
      return;
    }

    try {
      setParserTesting(true);
      setParserTestResult(null);

      const result = await identifyUniqueFromClipboard(
        database,
        parserTestText,
      );

      setParserTestResult(result);
    } catch (error) {
      console.error("Item parser test failed:", error);

      setParserTestResult({
        status: "invalid",
        parsed: {
          itemClass: null,
          rarity: null,
          name: null,
          baseType: null,
          canonicalName: null,
          canonicalBaseType: null,
          edition: "normal",
          rawText: parserTestText,
        },
        message:
          error instanceof Error ? error.message : String(error),
      });
    } finally {
      setParserTesting(false);
    }
  }

  async function handleCollectionProfileChange(
    profileId: string,
  ) {
    if (!database) {
      return;
    }

    const profileExists =
      collectionProfiles.some(
        (profile) => profile.id === profileId,
      );

    if (!profileExists) {
      return;
    }

    try {
      setAppError("");

      await withSqliteLockRetry(
  async () => {
    await database.execute(
      `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'active_collection_profile',
          ?
        )
      `,
      [profileId],
    );

    await loadCollectionData(
      database,
      profileId,
    );
  },
);

setActiveProfileId(
  profileId,
);
    } catch (error) {
      console.error(
        "Could not switch collection profile:",
        error,
      );

      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    }
  }

  async function handleSortModeChange(newSortMode: SortMode) {
    setSortMode(newSortMode);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('sort_mode', ?)
        `,
        [newSortMode],
      );
    } catch (error) {
      console.error("Could not save sort preference:", error);
    }
  }

  async function markUniqueMissing(
    uniqueId: string,
    profileId = activeProfileId,
  ) {
    if (!database) {
      return;
    }

    try {
      setAppError("");

      await database.execute(
        `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        [profileId, uniqueId],
      );

      await database.execute(
        `
          DELETE FROM profile_unique_tracking
          WHERE profile_id = ? AND unique_id = ?
        `,
        [profileId, uniqueId],
      );

      if (profileId === activeProfileId) {
        setUniques((current) =>
          current.map((item) =>
            item.id === uniqueId
              ? {
                  ...item,
                  flags: [],
                  reviewed: true,
                }
              : item,
          ),
        );
      }
    } catch (error) {
      console.error(error);

      setAppError(
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async function toggleTrackingFlag(
    uniqueId: string,
    flag: TrackingFlag,
  ) {
    if (!database) {
      return;
    }

    const unique = uniques.find(
      (item) => item.id === uniqueId,
    );

    if (!unique) {
      return;
    }

    const isActive =
      unique.flags.includes(flag);

    const enabledFlags: TrackingFlag[] =
      !isActive && flag === "foil"
        ? ["owned", "foil"]
        : [flag];

    try {
      setAppError("");

      await database.execute(
        `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        [activeProfileId, uniqueId],
      );

      if (isActive) {
        await database.execute(
          `
            DELETE FROM profile_unique_tracking
            WHERE
              profile_id = ?
              AND unique_id = ?
              AND flag = ?
          `,
          [
            activeProfileId,
            uniqueId,
            flag,
          ],
        );
      } else {
        for (const enabledFlag of enabledFlags) {
          await database.execute(
            `
              INSERT OR REPLACE INTO profile_unique_tracking (
                profile_id,
                unique_id,
                flag
              )
              VALUES (?, ?, ?)
            `,
            [
              activeProfileId,
              uniqueId,
              enabledFlag,
            ],
          );
        }
      }

      setUniques((current) =>
        current.map((item) => {
          if (item.id !== uniqueId) {
            return item;
          }

          return {
            ...item,
            reviewed: true,
            flags: isActive
              ? item.flags.filter(
                  (existingFlag) =>
                    existingFlag !== flag,
                )
              : Array.from(
                  new Set([
                    ...item.flags,
                    ...enabledFlags,
                  ]),
                ),
          };
        }),
      );
    } catch (error) {
      console.error(error);

      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    }
  }

  async function loadTrackableImportCatalogue(
    db: Database,
  ) {
    return db.select<
      {
        id: string;
        family_id: string;
        name: string;
        item_type: string;
        variant_label: string | null;
        release_version: string | null;
        is_legacy_only: number;
      }[]
    >(`
      SELECT
        v.id,
        v.family_id,
        v.name,
        v.item_type,
        v.variant_label,
        v.release_version,
        v.is_legacy_only
      FROM unique_variants v
      WHERE
        v.source IN ('poewiki', 'built-in-special')
        AND NOT (
          v.source = 'poewiki'
          AND EXISTS (
            SELECT 1
            FROM unique_variants special
            WHERE
              special.family_id = v.family_id
              AND special.source = 'built-in-special'
          )
        )
    `);
  }

  async function buildMissingOnlyPreview(
    db: Database,
    rows: ParsedImportRow[],
  ) {
    const catalogue =
      await loadTrackableImportCatalogue(db);

    const currentReleaseLine = latestReleaseLine(
      catalogue.map((item) => item.release_version),
    );

    const protectedLeagueUniques = currentReleaseLine
      ? catalogue
          .filter(
            (item) =>
              item.is_legacy_only !== 1 &&
              releaseLine(item.release_version) === currentReleaseLine &&
              !rows.some(
                (row) =>
                  row.itemType === item.item_type &&
                  isLikelySameImportName(row.name, item.name),
              ),
          )
          .map((item) => ({
            id: item.id,
            name: item.name,
            itemType: item.item_type,
            variantLabel: item.variant_label,
            releaseVersion: item.release_version,
          }))
          .sort((left, right) =>
            left.name.localeCompare(right.name) ||
            (left.variantLabel ?? "").localeCompare(
              right.variantLabel ?? "",
            ),
          )
      : [];

    return {
      latestReleaseLine: currentReleaseLine,
      protectedLeagueUniques,
    };
  }

  function buildImportedEntries(
    rows: ParsedImportRow[],
    mode: ImportMode,
  ): UniqueEntry[] {
    const rowsToImport =
      mode === "color-coded-list"
        ? rows.filter(
            (row) =>
              row.colorCodedStatus !==
              "unreviewed",
          )
        : rows;

    return rowsToImport.map((row) => {
      let normalizedStatus = "Missing";

      if (mode === "status-list") {
        const raw = row.rawStatus.trim();
        const normalized = raw.toLowerCase();

        if (normalized === "owned rainbow") {
          normalizedStatus = "Owned Foil";
        } else if (normalized === "owned foil") {
          normalizedStatus = "Owned Foil";
        } else if (normalized === "owned") {
          normalizedStatus = "Owned";
        } else if (normalized === "wearing") {
          normalizedStatus = "Wearing";
        } else if (normalized === "missing") {
          normalizedStatus = "Missing";
        } else {
          normalizedStatus = "Unreviewed";
        }
      } else if (mode === "color-coded-list") {
        if (
          row.colorCodedStatus ===
          "league-owned"
        ) {
          normalizedStatus =
            COLOR_CODED_IMPORT_STATUS.leagueOwned;
        } else if (
          row.colorCodedStatus ===
          "standard-owned"
        ) {
          normalizedStatus =
            COLOR_CODED_IMPORT_STATUS.standardOwned;
        } else {
          normalizedStatus =
            COLOR_CODED_IMPORT_STATUS.missing;
        }
      }

      return {
        id: row.id,
        familyId: makeFamilyId(row.name),
        name: row.name,
        baseType: null,
        itemType: row.itemType,
        variantLabel: null,
        importStatus: normalizedStatus,
        flags:
          mode === "color-coded-list"
            ? []
            : statusToFlags(
                normalizedStatus,
              ),
        editionAvailability: {
          ...DEFAULT_EDITION_AVAILABILITY,
        },
        editionSources: {
          ...DEFAULT_EDITION_SOURCES,
        },
        reviewed: true,
        isLegacyOnly: false,
        catalogueSource: "imported",
      };
    });
  }

  async function applyImportedMatchesToProfile(
    db: Database,
    destinationProfileId: string,
  ) {
    const matchedRows = await db.select<
      {
        canonical_variant_id: string;
        import_status: string;
      }[]
    >(`
      SELECT
        reconciliation.canonical_variant_id,
        imported.status AS import_status
      FROM import_reconciliation reconciliation
      JOIN imported_collection imported
        ON imported.id = reconciliation.import_id
      WHERE
        reconciliation.status = 'matched'
        AND reconciliation.canonical_variant_id IS NOT NULL
    `);

    const statements: SqliteTransactionStatement[] = [];

    for (const row of matchedRows) {
      const normalizedStatus = row.import_status.trim();
      const flags = statusToFlags(normalizedStatus);
      const isUnreviewed =
        normalizedStatus.toLowerCase() === "unreviewed";

      statements.push({
        sql: `
          DELETE FROM profile_unique_tracking
          WHERE profile_id = ? AND unique_id = ?
        `,
        params: [
          destinationProfileId,
          row.canonical_variant_id,
        ],
      });

      if (isUnreviewed) {
        statements.push({
          sql: `
            DELETE FROM profile_collection_review
            WHERE profile_id = ? AND unique_id = ?
          `,
          params: [
            destinationProfileId,
            row.canonical_variant_id,
          ],
        });
        continue;
      }

      statements.push({
        sql: `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        params: [
          destinationProfileId,
          row.canonical_variant_id,
        ],
      });

      for (const flag of flags) {
        statements.push({
          sql: `
            INSERT OR REPLACE INTO profile_unique_tracking (
              profile_id,
              unique_id,
              flag
            )
            VALUES (?, ?, ?)
          `,
          params: [
            destinationProfileId,
            row.canonical_variant_id,
            flag,
          ],
        });
      }
    }

    if (statements.length > 0) {
      await invoke("execute_sqlite_transaction", {
        statements,
        commit: true,
      });
    }
  }

  async function applyColorCodedImportToProfiles(
    db: Database,
    importedRows: ParsedImportRow[],
    leagueProfileId: string,
  ): Promise<ColorCodedImportSummary> {
    const activeProfiles =
      await db.select<
        {
          id: string;
          name: string;
          kind: string;
        }[]
      >(`
        SELECT
          id,
          name,
          kind
        FROM collection_profiles
        WHERE is_archived = 0
      `);

    const leagueProfile =
      activeProfiles.find(
        (profile) =>
          profile.id ===
            leagueProfileId &&
          profile.kind ===
            "challenge",
      );

    if (!leagueProfile) {
      throw new Error(
        "This color-coded list needs an active challenge league for the light-green entries.",
      );
    }

    const namedRows =
      importedRows.filter(
        (row) =>
          row.colorCodedStatus !==
          "unreviewed",
      );

    const reconciliationRows =
      await db.select<
        {
          import_id: string;
          canonical_variant_id:
            string | null;
          status: string;
        }[]
      >(`
        SELECT
          import_id,
          canonical_variant_id,
          status
        FROM import_reconciliation
      `);

    const matchedByImportId =
      new Map<string, string>();

    for (
      const row of
      reconciliationRows
    ) {
      if (
        row.status === "matched" &&
        row.canonical_variant_id
      ) {
        matchedByImportId.set(
          row.import_id,
          row.canonical_variant_id,
        );
      }
    }

    const matchedCanonicalIds =
      new Set(
        matchedByImportId.values(),
      );

    const statements:
      SqliteTransactionStatement[] =
      [];

    let standardOwned = 0;
    let leagueOwned = 0;
    let missing = 0;

    for (const row of namedRows) {
      const uniqueId =
        matchedByImportId.get(
          row.id,
        );

      if (!uniqueId) {
        continue;
      }

      /*
       * sheet describes the complete current state:
       * black = Missing everywhere,
       * light green = owned only in the selected league,
       * green-on-green = owned only in Standard.
       *
       * Archived league history is deliberately left alone.
       */
      statements.push(
        {
          sql: `
            DELETE FROM profile_unique_tracking
            WHERE
              unique_id = ?
              AND profile_id IN (
                SELECT id
                FROM collection_profiles
                WHERE is_archived = 0
              )
          `,
          params: [uniqueId],
        },
        {
          sql: `
            INSERT OR REPLACE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT
              id,
              ?,
              1
            FROM collection_profiles
            WHERE is_archived = 0
          `,
          params: [uniqueId],
        },
      );

      if (
        row.colorCodedStatus ===
        "standard-owned"
      ) {
        statements.push({
          sql: `
            INSERT OR REPLACE INTO profile_unique_tracking (
              profile_id,
              unique_id,
              flag
            )
            VALUES (?, ?, 'owned')
          `,
          params: [
            STANDARD_PROFILE_ID,
            uniqueId,
          ],
        });

        standardOwned += 1;
      } else if (
        row.colorCodedStatus ===
        "league-owned"
      ) {
        statements.push({
          sql: `
            INSERT OR REPLACE INTO profile_unique_tracking (
              profile_id,
              unique_id,
              flag
            )
            VALUES (?, ?, 'owned')
          `,
          params: [
            leagueProfileId,
            uniqueId,
          ],
        });

        leagueOwned += 1;
      } else {
        missing += 1;
      }
    }

    /*
     * "???" rows are deliberately unnamed by Path of Exile 2.
     * Use their position between surrounding resolved names to
     * identify the catalogue candidates that occupy that gap.
     *
     * We never guess ownership for those candidates: every
     * candidate in the gap becomes Unreviewed across all active
     * profiles. If no catalogue candidate can be located, the
     * placeholder is reported but no collection data is changed.
     */
    const catalogue =
      await loadTrackableImportCatalogue(
        db,
      );

    const rowsBySheet =
      new Map<
        string,
        ParsedImportRow[]
      >();

    for (const row of importedRows) {
      const existing =
        rowsBySheet.get(
          row.sourceSheetName,
        ) ?? [];

      existing.push(row);

      rowsBySheet.set(
        row.sourceSheetName,
        existing,
      );
    }

    const unknownCandidateIds =
      new Set<string>();

    let unresolvedPlaceholders = 0;

    for (
      const sheetRows of
      rowsBySheet.values()
    ) {
      const orderedRows =
        [...sheetRows].sort(
          (left, right) =>
            left.sourceRowIndex -
            right.sourceRowIndex,
        );

      const itemType =
        orderedRows[0]?.itemType;

      if (!itemType) {
        continue;
      }

      const typeCatalogue =
        catalogue
          .filter(
            (item) =>
              item.item_type ===
                itemType &&
              item.is_legacy_only !== 1,
          )
          .sort(
            (left, right) =>
              importNameWithAlias(
                left.name,
              ).localeCompare(
                importNameWithAlias(
                  right.name,
                ),
              ) ||
              (
                left.variant_label ?? ""
              ).localeCompare(
                right.variant_label ??
                  "",
              ),
          );

      const catalogueIndexById =
        new Map(
          typeCatalogue.map(
            (item, index) =>
              [item.id, index] as const,
          ),
        );

      let rowIndex = 0;

      while (
        rowIndex <
        orderedRows.length
      ) {
        if (
          orderedRows[rowIndex]
            .colorCodedStatus !==
          "unreviewed"
        ) {
          rowIndex += 1;
          continue;
        }

        const runStart =
          rowIndex;

        while (
          rowIndex <
            orderedRows.length &&
          orderedRows[rowIndex]
            .colorCodedStatus ===
            "unreviewed"
        ) {
          rowIndex += 1;
        }

        const runEnd =
          rowIndex;

        let previousCatalogueIndex:
          | number
          | null = null;

        for (
          let previous =
            runStart - 1;
          previous >= 0;
          previous -= 1
        ) {
          const previousId =
            matchedByImportId.get(
              orderedRows[
                previous
              ].id,
            );

          const previousIndex =
            previousId
              ? catalogueIndexById.get(
                  previousId,
                )
              : undefined;

          if (
            previousIndex !==
            undefined
          ) {
            previousCatalogueIndex =
              previousIndex;
            break;
          }
        }

        let nextCatalogueIndex:
          | number
          | null = null;

        for (
          let next = runEnd;
          next <
          orderedRows.length;
          next += 1
        ) {
          const nextId =
            matchedByImportId.get(
              orderedRows[next].id,
            );

          const nextIndex =
            nextId
              ? catalogueIndexById.get(
                  nextId,
                )
              : undefined;

          if (
            nextIndex !==
            undefined
          ) {
            nextCatalogueIndex =
              nextIndex;
            break;
          }
        }

        const startIndex =
          previousCatalogueIndex ===
          null
            ? 0
            : previousCatalogueIndex +
              1;

        const endIndex =
          nextCatalogueIndex === null
            ? typeCatalogue.length
            : nextCatalogueIndex;

        const candidates =
          typeCatalogue
            .slice(
              startIndex,
              endIndex,
            )
            .filter(
              (item) =>
                !matchedCanonicalIds.has(
                  item.id,
                ),
            );

        if (
          candidates.length === 0
        ) {
          unresolvedPlaceholders +=
            runEnd -
            runStart;

          continue;
        }

        for (
          const candidate of
          candidates
        ) {
          unknownCandidateIds.add(
            candidate.id,
          );
        }
      }
    }

    for (
      const uniqueId of
      unknownCandidateIds
    ) {
      statements.push(
        {
          sql: `
            DELETE FROM profile_unique_tracking
            WHERE
              unique_id = ?
              AND profile_id IN (
                SELECT id
                FROM collection_profiles
                WHERE is_archived = 0
              )
          `,
          params: [uniqueId],
        },
        {
          sql: `
            DELETE FROM profile_collection_review
            WHERE
              unique_id = ?
              AND profile_id IN (
                SELECT id
                FROM collection_profiles
                WHERE is_archived = 0
              )
          `,
          params: [uniqueId],
        },
      );
    }

    if (statements.length > 0) {
      await invoke(
        "execute_sqlite_transaction",
        {
          statements,
          commit: true,
        },
      );
    }

    const unresolvedNamedRows =
      namedRows.filter(
        (row) =>
          !matchedByImportId.has(
            row.id,
          ),
      ).length;

    return {
      kind: "color-coded-list",
      standardOwned,
      leagueOwned,
      missing,
      unknownCandidatesMarkedUnreviewed:
        unknownCandidateIds.size,
      unresolvedPlaceholders,
      unresolvedNamedRows,
      leagueProfileId,
      leagueProfileName:
        leagueProfile.name,
    };
  }

  async function applyMissingOnlyInferenceToProfile(
    db: Database,
    importedRows: ParsedImportRow[],
    destinationProfileId: string,
  ): Promise<MissingOnlyInferenceSummary> {
    const catalogue =
      await loadTrackableImportCatalogue(db);

    const currentReleaseLine = latestReleaseLine(
      catalogue.map((item) => item.release_version),
    );

    const familyCounts = new Map<string, number>();

    for (const item of catalogue) {
      familyCounts.set(
        item.family_id,
        (familyCounts.get(item.family_id) ?? 0) + 1,
      );
    }

    const reconciliationRows = await db.select<
      {
        import_id: string;
        canonical_variant_id: string | null;
        status: string;
        candidates_json: string;
        name: string;
        item_type: string;
      }[]
    >(`
      SELECT
        reconciliation.import_id,
        reconciliation.canonical_variant_id,
        reconciliation.status,
        reconciliation.candidates_json,
        imported.name,
        imported.item_type
      FROM import_reconciliation reconciliation
      JOIN imported_collection imported
        ON imported.id = reconciliation.import_id
    `);

    const explicitMissingIds = new Set<string>();
    const protectedIds = new Set<string>();

    for (const row of reconciliationRows) {
      if (
        row.status === "matched" &&
        row.canonical_variant_id
      ) {
        explicitMissingIds.add(
          row.canonical_variant_id,
        );
        continue;
      }

      if (row.status === "ambiguous") {
        try {
          const candidates = JSON.parse(
            row.candidates_json,
          ) as Array<{ id?: string }>;

          for (const candidate of candidates) {
            if (candidate.id) {
              protectedIds.add(candidate.id);
            }
          }
        } catch {
          // An unreadable candidate list should never make us infer ownership.
        }
      }

      if (row.status === "unmatched") {
        for (const item of catalogue) {
          if (
            item.item_type === row.item_type &&
            isLikelySameImportName(
              row.name,
              item.name,
            )
          ) {
            protectedIds.add(item.id);
          }
        }
      }
    }

    for (const item of catalogue) {
      if ((familyCounts.get(item.family_id) ?? 0) > 1) {
        protectedIds.add(item.id);
      }

      // If release metadata is missing, absence is not enough evidence to
      // infer ownership safely.
      if (!releaseLine(item.release_version)) {
        protectedIds.add(item.id);
      }
    }

    const protectedNewLeague = currentReleaseLine
      ? catalogue
          .filter(
            (item) =>
              item.is_legacy_only !== 1 &&
              releaseLine(item.release_version) === currentReleaseLine &&
              !explicitMissingIds.has(item.id) &&
              !importedRows.some(
                (row) =>
                  row.itemType === item.item_type &&
                  isLikelySameImportName(
                    row.name,
                    item.name,
                  ),
              ),
          )
          .map((item) => ({
            id: item.id,
            name: item.name,
            itemType: item.item_type,
            variantLabel: item.variant_label,
            releaseVersion: item.release_version,
          }))
          .sort((left, right) =>
            left.name.localeCompare(right.name) ||
            (left.variantLabel ?? "").localeCompare(
              right.variantLabel ?? "",
            ),
          )
      : [];

    const protectedNewIds = new Set(
      protectedNewLeague.map((item) => item.id),
    );

    const inferredOwnedIds = currentReleaseLine
      ? catalogue
          .filter(
            (item) =>
              item.is_legacy_only !== 1 &&
              !explicitMissingIds.has(item.id) &&
              !protectedIds.has(item.id) &&
              !protectedNewIds.has(item.id) &&
              releaseLine(item.release_version) !== null &&
              releaseLine(item.release_version) !== currentReleaseLine,
          )
          .map((item) => item.id)
      : [];

    const statements: SqliteTransactionStatement[] = [];
    const CHUNK_SIZE = 250;

    for (
      let index = 0;
      index < explicitMissingIds.size;
      index += CHUNK_SIZE
    ) {
      const ids = Array.from(explicitMissingIds).slice(
        index,
        index + CHUNK_SIZE,
      );
      const placeholders = ids.map(() => "?").join(", ");

      statements.push(
        {
          sql: `
            DELETE FROM profile_unique_tracking
            WHERE
              profile_id = ?
              AND unique_id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
        {
          sql: `
            INSERT OR REPLACE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT ?, id, 1
            FROM unique_variants
            WHERE id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
      );
    }

    for (
      let index = 0;
      index < inferredOwnedIds.length;
      index += CHUNK_SIZE
    ) {
      const ids = inferredOwnedIds.slice(
        index,
        index + CHUNK_SIZE,
      );
      const placeholders = ids.map(() => "?").join(", ");

      statements.push(
        {
          sql: `
            INSERT OR REPLACE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT ?, id, 1
            FROM unique_variants
            WHERE id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
        {
          sql: `
            INSERT OR IGNORE INTO profile_unique_tracking (
              profile_id,
              unique_id,
              flag
            )
            SELECT ?, id, 'owned'
            FROM unique_variants
            WHERE id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
      );
    }

    if (statements.length > 0) {
      await invoke("execute_sqlite_transaction", {
        statements,
        commit: true,
      });
    }

    const destinationProfileName =
      collectionProfiles.find(
        (profile) => profile.id === destinationProfileId,
      )?.name ?? "Selected Collection";

    return {
      kind: "missing-only",
      inferredOwned: inferredOwnedIds.length,
      explicitMissing: explicitMissingIds.size,
      protectedNewLeague,
      protectedAmbiguous: protectedIds.size,
      latestReleaseLine: currentReleaseLine,
      destinationProfileId,
      destinationProfileName,
    };
  }

  function openImportReviewScreen() {
    setSettingsOpen(false);

    if (
      sourceImportMode ===
      "color-coded-list"
    ) {
      setAppError(
        "Manual row matching is disabled for this color-coded import because one spreadsheet row can affect Standard and a challenge league differently.",
      );
      return;
    }

    setBatchImportReviewOpen(true);
  }

  async function saveCollection(
    imported: UniqueEntry[],
    fileName: string,
    mode: ImportMode,
    parsedRows: ParsedImportRow[],
    destinationProfileId: string,
  ): Promise<ImportApplicationSummary | null> {
    if (!database) {
      throw new Error(
        "The local database is not ready yet.",
      );
    }

    const previousImports = await database.select<
      { id: string }[]
    >(`
      SELECT id
      FROM imported_collection
    `);

    for (const previous of previousImports) {
      // These IDs belong to raw import rows, not the canonical catalogue.
      // Canonical tracking uses stable poewiki:/special: IDs.
      await database.execute(
        `
          DELETE FROM unique_tracking
          WHERE unique_id = ?
            AND unique_id NOT IN (
              SELECT id
              FROM unique_variants
              WHERE source IN ('poewiki', 'built-in-special')
            )
        `,
        [previous.id],
      );

      await database.execute(
        `
          DELETE FROM collection_review
          WHERE unique_id = ?
            AND unique_id NOT IN (
              SELECT id
              FROM unique_variants
              WHERE source IN ('poewiki', 'built-in-special')
            )
        `,
        [previous.id],
      );
    }

    await database.execute(
      "DELETE FROM imported_collection",
    );

    await database.execute(`
      DELETE FROM import_reconciliation
    `).catch(() => undefined);

    await database.execute(`
      DELETE FROM import_review_progress
    `).catch(() => undefined);

    await database.execute(`
      DELETE FROM import_manual_resolutions
    `).catch(() => undefined);

    for (const unique of imported) {
      await database.execute(
        `
          INSERT INTO imported_collection (
            id,
            name,
            item_type,
            status
          )
          VALUES (?, ?, ?, ?)
        `,
        [
          unique.id,
          unique.name,
          unique.itemType,
          unique.importStatus,
        ],
      );

      // Raw spreadsheet rows are reference data. Reconciliation copies their
      // state onto canonical IDs where a safe match exists.
      await database.execute(
        `
          INSERT OR REPLACE INTO collection_review (
            unique_id,
            reviewed
          )
          VALUES (?, 1)
        `,
        [unique.id],
      );

      for (const flag of unique.flags) {
        await database.execute(
          `
            INSERT OR REPLACE INTO unique_tracking (
              unique_id,
              flag
            )
            VALUES (?, ?)
          `,
          [unique.id, flag],
        );
      }
    }

    await database.execute(
      `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES ('source_file', ?)
      `,
      [fileName],
    );

    await database.execute(
      `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES ('source_import_mode', ?)
      `,
      [mode],
    );

    setSourceImportMode(mode);

    const canonicalReady =
      await isCanonicalCatalogueReady(database);

    if (!canonicalReady) {
      if (
        mode === "missing-only" ||
        mode === "color-coded-list"
      ) {
        throw new Error(
          mode === "color-coded-list"
            ? "The color-coded import needs the canonical catalogue before the tracker can safely match names and ??? placeholders."
            : "Missing-only imports need the canonical catalogue before ownership can be inferred.",
        );
      }

      return null;
    }

    const reconciliation =
      await reconcileImportedCollection(
        database,
      );

    setImportMatchSummary(reconciliation);

    if (mode === "color-coded-list") {
      return applyColorCodedImportToProfiles(
        database,
        parsedRows,
        destinationProfileId,
      );
    }

    await applyImportedMatchesToProfile(
      database,
      destinationProfileId,
    );

    if (mode === "missing-only") {
      return applyMissingOnlyInferenceToProfile(
        database,
        parsedRows,
        destinationProfileId,
      );
    }

    return null;
  }

  async function confirmPendingImport() {
    if (!pendingImport || !database || isImporting) {
      return;
    }

    try {
      setIsImporting(true);
      setAppError("");

      if (
        pendingImportMode ===
          "color-coded-list" &&
        pendingImportProfileId ===
          STANDARD_PROFILE_ID
      ) {
        throw new Error(
          "Choose the challenge league that the light-green entries belong to.",
        );
      }

      const imported = buildImportedEntries(
        pendingImport.rows,
        pendingImportMode,
      );

      const summary = await saveCollection(
        imported,
        pendingImport.fileName,
        pendingImportMode,
        pendingImport.rows,
        pendingImportProfileId,
      );

      if (
        pendingImportMode ===
          "color-coded-list" ||
        pendingImportProfileId ===
          activeProfileId
      ) {
        await loadCollectionData(
          database,
          activeProfileId,
        );
      }

      setSourceFile(
        pendingImport.fileName,
      );
      setPendingImport(null);
      setImportDetailsExpanded(false);

      if (
        summary?.kind ===
        "missing-only"
      ) {
        setMissingOnlySummary(
          summary,
        );
      } else if (
        summary?.kind ===
        "color-coded-list"
      ) {
        setColorCodedSummary(
          summary,
        );
      }
    } catch (error) {
      console.error(error);
      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    } finally {
      setIsImporting(false);
    }
  }

  async function handleImport() {
    if (!database) {
      setAppError("The local database is not ready yet.");
      return;
    }

    try {
      setAppError("");
      setMissingOnlySummary(null);
      setColorCodedSummary(null);
      setImportDetailsExpanded(false);
      setOtherImportModesOpen(false);

      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx", "xls"],
          },
        ],
      });

      if (!selected || Array.isArray(selected)) {
        return;
      }

      setIsImporting(true);

      const fileBytes = await readFile(selected);

      const workbook = read(fileBytes, {
        type: "array",
        cellFormula: false,
        cellHTML: false,
        cellStyles: true,
      });

      const parsedRows: ParsedImportRow[] = [];
      let missingOnlyScore = 0;
      let statusListScore = 0;
      let oleColorEvidence = 0;

      for (const sheetName of workbook.SheetNames) {
        const itemType = normalizeImportItemType(sheetName);

        // Auxiliary sheets such as ALT sheet are intentionally ignored.
        // Only sheets that map to a real PoE 2 Unique Tracker item type are imported.
        if (!itemType) {
          continue;
        }

        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
          continue;
        }

        const rows = utils.sheet_to_json<
          (string | number | boolean)[]
        >(sheet, {
          header: 1,
          defval: "",
          raw: false,
        });

        rows.forEach((row, rowIndex) => {
          const name = String(
            row[0] ?? "",
          ).trim();

          const secondColumn =
            String(
              row[1] ?? "",
            ).trim();

          const thirdColumn =
            String(
              row[2] ?? "",
            ).trim();

          if (rowIndex === 0) {
            const firstHeader =
              normalizeImportName(
                name,
              ).replace(/:$/, "");

            const secondHeader =
              normalizeImportName(
                secondColumn,
              ).replace(/:$/, "");

            const thirdHeader =
              normalizeImportName(
                thirdColumn,
              ).replace(/:$/, "");

            if (
              secondHeader.includes(
                "price",
              ) ||
              thirdHeader.includes(
                "acquisition",
              )
            ) {
              missingOnlyScore += 4;
            }

            if (
              secondHeader.includes(
                "status",
              )
            ) {
              statusListScore += 4;
            }

            if (
              firstHeader === "name" ||
              firstHeader ===
                "item name"
            ) {
              return;
            }
          }

          if (!name) {
            return;
          }

          const firstColumnCell =
            sheet[
              utils.encode_cell({
                r: rowIndex,
                c: 0,
              })
            ] as
              | SheetJsStyledCell
              | undefined;

          const colorCodedStatus =
            getColorCodedStatus(
              firstColumnCell,
              name,
            );

          if (
            colorCodedStatus ===
              "standard-owned" ||
            colorCodedStatus ===
              "league-owned"
          ) {
            oleColorEvidence += 3;
          } else if (
            colorCodedStatus ===
            "unreviewed"
          ) {
            oleColorEvidence += 1;
          }

          if (
            KNOWN_IMPORT_STATUSES.has(
              secondColumn.toLowerCase(),
            )
          ) {
            statusListScore += 2;
          }

          parsedRows.push({
            id:
              `${sheetName}-${rowIndex}-${name}`,
            name,
            itemType,
            rawStatus: secondColumn,
            sourceSheetName:
              sheetName,
            sourceRowIndex:
              rowIndex,
            colorCodedStatus,
          });
        });
      }

      if (parsedRows.length === 0) {
        throw new Error(
          "No recognizable unique-list rows were found in that workbook.",
        );
      }

      const fileName =
        selected.split(/[\\/]/).pop() ?? selected;

      const preview = await buildMissingOnlyPreview(
        database,
        parsedRows,
      );

      const suggestedMode: ImportMode =
        oleColorEvidence >= 3
          ? "color-coded-list"
          : missingOnlyScore >
              statusListScore
            ? "missing-only"
            : "status-list";

      setPendingImport({
        fileName,
        rows: parsedRows,
        suggestedMode,
        latestReleaseLine:
          preview.latestReleaseLine,
        protectedLeagueUniques:
          preview.protectedLeagueUniques,
      });

      setPendingImportMode(
        suggestedMode,
      );

      setOtherImportModesOpen(
        suggestedMode ===
          "color-coded-list",
      );

      const activeChallengeProfile =
        collectionProfiles.find(
          (profile) =>
            profile.id ===
              activeProfileId &&
            profile.kind ===
              "challenge",
        );

      const firstChallengeProfile =
        collectionProfiles.find(
          (profile) =>
            profile.kind ===
            "challenge",
        );

      setPendingImportProfileId(
        suggestedMode ===
          "color-coded-list"
          ? (
              activeChallengeProfile ??
              firstChallengeProfile
            )?.id ??
              STANDARD_PROFILE_ID
          : collectionProfiles.some(
                (profile) =>
                  profile.id ===
                  activeProfileId,
              )
            ? activeProfileId
            : STANDARD_PROFILE_ID,
      );
    } catch (error) {
      console.error(error);

      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
        } finally {
      setIsImporting(false);
    }
  }

  async function minimizeMainWindow() {
    setClosePromptOpen(false);

    const mainWindow =
      await WebviewWindow.getByLabel(
        "main",
      );

    if (mainWindow) {
      await mainWindow.minimize();
    }
  }

  async function hideMainWindowToTray() {
    setClosePromptOpen(false);

    const mainWindow =
      await WebviewWindow.getByLabel(
        "main",
      );

    if (mainWindow) {
      await mainWindow.hide();
    }
  }

  return (
  <main className="app">
    <AppUpdater />

    <header className="app-header">
        <div>
          <h1>PoE 2 Unique Tracker</h1>
          <p className="subtitle">
            Path of Exile 2 Unique Collection Tracker
          </p>
        </div>

        <button
          className="settings-button"
          type="button"
          onClick={() => setSettingsOpen(true)}
        >
          Settings
        </button>
      </header>

      <section className="collection-card">
        <div className="collection-info">
          <span className="label">Active Collection</span>

          <select
            className="collection-profile-select"
            value={activeProfileId}
            disabled={!databaseReady}
            onChange={(event) =>
              void handleCollectionProfileChange(
                event.target.value,
              )
            }
          >
            {collectionProfiles.map((profile) => (
              <option
                key={profile.id}
                value={profile.id}
              >
                {profile.name}
              </option>
            ))}
          </select>

          {rolloverPreview &&
  !rolloverPreviewOpen &&
  (rolloverMode === "pending" ||
    rolloverMode === "dev-pending") && (
    <button
      type="button"
      className="catalogue-update-secondary"
      style={{
        alignSelf: "flex-start",
        marginTop: "6px",
      }}
      onClick={() => {
        setRolloverChangesExpanded(false);
        setRolloverPreviewOpen(true);
      }}
    >
      {"\u26A0"} {rolloverPreview.oldLeagueName} appears to have ended
      {" \u2014 "}
      Review
    </button>
  )}

          <span className="count">
            {databaseReady
              ? `${uniques.length.toLocaleString()} catalogue entries loaded`
              : "Starting local database..."}
          </span>

          {activeProfile?.kind === "challenge" && (
            <span className="source-file">
              Fresh challenge-league collection
            </span>
          )}

          {activeProfile?.kind === "standard" &&
            sourceFile && (
              <span className="source-file">
                Imported from {sourceFile}
              </span>
            )}
        </div>

        <button
          className="import-button"
          type="button"
          onClick={handleImport}
          disabled={isImporting || !databaseReady}
        >
          {isImporting
            ? "Importing..."
            : uniques.length > 0
              ? "Import Another Collection"
              : "Import Existing Collection"}
        </button>
      </section>

      {appError && (
        <section className="error-message">
          <strong>Something went wrong</strong>
          <span>{appError}</span>
        </section>
      )}

      {uniques.length === 0 ? (
        <section className="empty-state">
          <h2>No collection loaded yet</h2>
          <p>
            Import your existing spreadsheet to start tracking your Path of
            Exile uniques.
          </p>
        </section>
      ) : (
        <>
          <section className="status-summary">
            {trackingCounts.map(({ label, count }) => (
              <div className="status-card" key={label}>
                <span>{label}</span>
                <strong>{count.toLocaleString()}</strong>
              </div>
            ))}
          </section>

          <section className="filter-panel">
            <div className="search-field">
  <label htmlFor="unique-search">Search</label>

  <div className="search-input-wrap">
    <input
      id="unique-search"
      type="text"
      placeholder="Search uniques..."
      value={searchTerm}
      onChange={(event) =>
        setSearchTerm(event.target.value)
      }
    />

    {searchTerm && (
      <button
        type="button"
        className="search-clear-button"
        aria-label="Clear search"
        title="Clear search"
        onClick={() => setSearchTerm("")}
      >
        {"\u00D7"}
      </button>
    )}
  </div>
</div>

            <div className="filter-field">
              <label htmlFor="sort-mode">Sort</label>

              <select
                id="sort-mode"
                value={sortMode}
                onChange={(event) =>
                  handleSortModeChange(event.target.value as SortMode)
                }
              >
                <option value="alphabetical">{"All A\u2013Z"}</option>
                <option value="type">{"By Type \u2192 A\u2013Z"}</option>
              </select>
            </div>

            <div className="filter-field">
              <label htmlFor="status-filter">Status</label>

              <select
                id="status-filter"
                value={statusFilter}
                onChange={(event) =>
                  setStatusFilter(event.target.value as StatusFilter)
                }
              >
                <option value="all">All Statuses</option>
                <option value="unreviewed">Unreviewed</option>
                <option value="missing">Missing</option>
                <option value="owned">Owned</option>
                <option value="wearing">Wearing</option>

                {extraTracking.foil && (
                  <option value="foil">Foil</option>
                )}

                {extraTracking.foulborn && (
                  <option value="foulborn">Foulborn</option>
                )}

                {extraTracking.vestigial && (
                  <option value="vestigial">Vestigial</option>
                )}
              </select>
            </div>

            <div className="filter-field">
              <label htmlFor="type-filter">Type</label>

              <select
                id="type-filter"
                value={typeFilter}
                onChange={(event) => setTypeFilter(event.target.value)}
              >
                <option value="All">All Types</option>

                {availableTypes.map((itemType) => (
                  <option value={itemType} key={itemType}>
                    {itemType}
                  </option>
                ))}
              </select>
            </div>
            <button
  type="button"
  className="clear-filters-button"
  disabled={
    searchTerm === "" &&
    statusFilter === "all" &&
    typeFilter === "All"
  }
  onClick={() => {
    setSearchTerm("");
    setStatusFilter("all");
    setTypeFilter("All");
  }}
>
  Clear Filters
</button>
          </section>

          <section className="collection-list">
            <div className="list-header">
              <div>
                <h2>Collection</h2>
                <p>
                  Showing {displayedUniques.length.toLocaleString()} of{" "}
                  {uniques.length.toLocaleString()} catalogue entries.
                </p>
              </div>

              <span>
                {displayedUniques.length.toLocaleString()} shown
              </span>
            </div>

            <div className="unique-table">
              {displayedUniques.length === 0 ? (
                <div className="no-results">
                  No uniques match your current filters.
                </div>
              ) : sortMode === "alphabetical" ? (
                <>
                  <div className="unique-row table-heading">
                    <span>Name</span>
                    <span>Type</span>
                    <span>Tracking</span>
                  </div>

                  {displayedUniques.map((unique) => (
                    <div className="unique-row" key={unique.id}>
                      <span className="unique-name-cell">
                        <span className="unique-name-stack">
                          <span className="unique-name">{unique.name}</span>
                          {(unique.baseType || unique.variantLabel) && (
                            <span className="unique-variant-meta">
                              {[unique.baseType, unique.variantLabel]
                                .filter(Boolean)
                                .join(" \u2022 ")}
                            </span>
                          )}
                        </span>
                        {!unique.reviewed && (
                          <span className="unreviewed-badge">
                            NEW / UNREVIEWED
                          </span>
                        )}
                        {unique.isLegacyOnly && (
                          <span className="legacy-badge">
                            LEGACY
                          </span>
                        )}
                      </span>
                      <span className="unique-type">{unique.itemType}</span>

                      <TrackingBadges
                        flags={unique.flags}
                        colors={statusColors}
                        extraTracking={extraTracking}
editionAvailability={
  unique.editionAvailability
}

editionSources={
  unique.editionSources
}
                        isCollected={isUniqueCollected(unique)}
                        isMissing={isUniqueMissing(unique)}
                        onToggle={(flag) =>
                          toggleTrackingFlag(unique.id, flag)
                        }
                        onMarkMissing={() =>
                          markUniqueMissing(unique.id)
                        }
                      />
                    </div>
                  ))}
                </>
              ) : (
                groupedUniques.map((group) => (
                  <section className="type-group" key={group.itemType}>
                    <div className="type-group-header">
                      <span className="type-group-header-label">
                        {group.itemType.toUpperCase()}
                      </span>
                    </div>

                    {group.items.map((unique) => (
                      <div
                        className="unique-row grouped-row"
                        key={unique.id}
                      >
                        <span className="unique-name-cell">
                          <span className="unique-name-stack">
                            <span className="unique-name">{unique.name}</span>
                            {(unique.baseType || unique.variantLabel) && (
                              <span className="unique-variant-meta">
                                {[unique.baseType, unique.variantLabel]
                                  .filter(Boolean)
                                  .join(" \u2022 ")}
                              </span>
                            )}
                          </span>
                          {!unique.reviewed && (
                            <span className="unreviewed-badge">
                              NEW / UNREVIEWED
                            </span>
                          )}
                          {unique.isLegacyOnly && (
                            <span className="legacy-badge">
                              LEGACY
                            </span>
                          )}
                        </span>

                        <TrackingBadges
                          flags={unique.flags}
                          colors={statusColors}
                          extraTracking={extraTracking}
editionAvailability={
  unique.editionAvailability
}

editionSources={
  unique.editionSources
}
                          isCollected={isUniqueCollected(unique)}
                          isMissing={isUniqueMissing(unique)}
                          onToggle={(flag) =>
                            toggleTrackingFlag(unique.id, flag)
                          }
                          onMarkMissing={() =>
                            markUniqueMissing(unique.id)
                          }
                        />
                      </div>
                    ))}
                  </section>
                ))
              )}
            </div>
                    </section>

          {hasUncertainEditionData && (
            <p className="edition-data-note">
              *Data inconclusive. Starred editions
              may or may not exist.
            </p>
          )}
        </>
      )}

      {rolloverPreviewOpen &&
  rolloverPreview && (
    <div className="catalogue-update-overlay">
      <section className="catalogue-update-modal">
        <div className="catalogue-update-heading">
          <span className="catalogue-update-kicker">
            {rolloverMode === "dev-pending"
              ? "DEV LEAGUE ROLLOVER TEST"
              : rolloverMode === "dev-complete"
                ? "DEV ROLLOVER TEST PASSED"
                : rolloverMode === "pending"
                  ? "LEAGUE ENDED"
                  : "LEAGUE ROLLOVER COMPLETE"}
          </span>

          <h2>
            {rolloverPreview.oldLeagueName}
            {" \u2192 "}
            {rolloverPreview.newLeagueName}
          </h2>

          <p>
            {rolloverMode === "dev-pending"
              ? "Development simulation. This tests archiving one league and merging its tracked collection into Standard, then rolls every database change back."
              : rolloverMode === "dev-complete"
                ? "The complete rollover transaction succeeded and was rolled back. No collection data was changed."
                : rolloverMode === "pending"
                  ? `${rolloverPreview.oldLeagueName} no longer appears in the active league list. Its tracked collection can now be archived and merged into Standard.`
                  : `${rolloverPreview.oldLeagueName} has been archived and merged into Standard. Other active league collections were left untouched.`}
          </p>
        </div>

        <div className="catalogue-update-stats">
          <div>
            <strong>
              {rolloverPreview.changedUniques}
            </strong>
            <span>uniques changed</span>
          </div>

          <div>
            <strong>
              {rolloverPreview.flagCounts.owned}
            </strong>
            <span>Owned added</span>
          </div>

          <div>
            <strong>
              {rolloverPreview.flagCounts.wearing}
            </strong>
            <span>Wearing added</span>
          </div>

          <div>
            <strong>
              {rolloverPreview.flagCounts.foil}
            </strong>
            <span>Foils added</span>
          </div>
        </div>



        {rolloverChangesExpanded &&
          rolloverPreview.changes.length > 0 && (
            <div
              style={{
                maxHeight: "220px",
                overflowY: "auto",
              }}
            >
              {rolloverPreview.changes.map(
                (change) => (
                  <p
                    className="catalogue-update-note"
                    key={change.uniqueId}
                  >
                    <strong>
                      {change.name}
                    </strong>
                    {" \u2014 "}
                    {change.addedFlags
                      .map(
                        (flag) =>
                          STATUS_LABELS[flag],
                      )
                      .join(", ")}
                  </p>
                ),
              )}
            </div>
          )}

        <div className="catalogue-update-actions">
          {rolloverPreview.changes.length > 0 && (
            <button
              type="button"
              className="catalogue-update-secondary"
              onClick={() =>
                setRolloverChangesExpanded(
                  (current) => !current,
                )
              }
            >
              {rolloverChangesExpanded
                ? "Hide Standard Changes"
                : "View Standard Changes"}
            </button>
          )}

          {(rolloverMode === "pending" ||
            rolloverMode === "dev-pending") && (
            <button
              type="button"
              className="catalogue-update-secondary"
              disabled={rolloverApplying}
              onClick={() =>
                setRolloverPreviewOpen(false)
              }
            >
              Not Now
            </button>
          )}

          {rolloverMode === "pending" ? (
            <button
              type="button"
              className="catalogue-update-primary"
              disabled={rolloverApplying}
              onClick={() =>
                void confirmLeagueRollover()
              }
            >
              {rolloverApplying
                ? "Migrating League..."
                : "Merge into Standard"}
            </button>
          ) : rolloverMode ===
            "dev-pending" ? (
            <button
              type="button"
              className="catalogue-update-primary"
              disabled={rolloverApplying}
              onClick={() =>
                void confirmDevLeagueRollover()
              }
            >
              {rolloverApplying
                ? "Testing Rollover..."
                : "Run Rollover Test"}
            </button>
          ) : (
            <button
              type="button"
              className="catalogue-update-primary"
              onClick={() =>
                setRolloverPreviewOpen(false)
              }
            >
              {rolloverMode === "dev-complete"
                ? "Close Test"
                : "Continue"}
            </button>
          )}
        </div>
      </section>
    </div>
  )}

      {activeDiscoveryNotice && (
        <div className="catalogue-update-overlay">
          <section
            className="catalogue-update-modal"
            style={{
              width:
                "min(720px, calc(100vw - 40px))",
              maxWidth: 720,
            }}
          >
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                {activeDiscoveryNotice.kind ===
                "new-league"
                  ? "NEW LEAGUE DETECTED"
                  : activeDiscoveryNotice.items.length ===
                      1
                    ? "NEW UNIQUE ADDED TO TRACKER"
                    : "NEW UNIQUES ADDED TO TRACKER"}
              </span>

              <h2>
                {activeDiscoveryNotice.kind ===
                "new-league"
                  ? activeDiscoveryNotice.leagueName
                  : activeDiscoveryNotice.items.length ===
                      1
                    ? activeDiscoveryNotice.items[0]?.name ??
                      "New unique"
                    : `${activeDiscoveryNotice.items.length} new uniques found`}
              </h2>

              <p>
                {activeDiscoveryNotice.kind ===
                "new-league"
                  ? `${activeDiscoveryNotice.leagueName} is now available as a fresh collection. ${activeDiscoveryNotice.items.length} new ${
                      activeDiscoveryNotice.items.length === 1
                        ? "unique was"
                        : "uniques were"
                    } added to the tracker during the league refresh.`
                  : `${activeDiscoveryNotice.items.length} new ${
                      activeDiscoveryNotice.items.length === 1
                        ? "unique was"
                        : "uniques were"
                    } found by the daily catalogue check. ${
                      activeDiscoveryNotice.items.length === 1
                        ? "It is"
                        : "They are"
                    } marked Unreviewed until you choose a collection status.`}
              </p>
            </div>

            <div className="catalogue-update-stats">
              <div>
                <strong>
                  {activeDiscoveryNotice.items.length}
                </strong>
                <span>
                  {activeDiscoveryNotice.items.length ===
                  1
                    ? "new unique"
                    : "new uniques"}
                </span>
              </div>

              <div>
                <strong>
                  {activeDiscoveryNotice.kind ===
                  "new-league"
                    ? "Missing"
                    : "Unreviewed"}
                </strong>
                <span>starting status</span>
              </div>
            </div>

            {activeDiscoveryNotice.items.length >
              0 && (
              <>
                <div className="catalogue-update-actions">
                  <button
                    type="button"
                    className="catalogue-update-secondary"
                    onClick={() =>
                      setDiscoveryDetailsExpanded(
                        (current) => !current,
                      )
                    }
                  >
                    {discoveryDetailsExpanded
                      ? "Hide New Uniques"
                      : "View New Uniques"}
                  </button>
                </div>

                {discoveryDetailsExpanded && (
                  <div
                    style={{
                      maxHeight: 280,
                      overflowY: "auto",
                      borderTop:
                        "1px solid #3c352d",
                    }}
                  >
                    {activeDiscoveryNotice.items.map(
                      (item) => (
                        <p
                          className="catalogue-update-note"
                          key={item.id}
                        >
                          <strong>
                            {item.name}
                          </strong>
                          {item.variantLabel
                            ? ` \u2014 ${item.variantLabel}`
                            : ""}
                          {" \u2022 "}
                          {item.itemType}
                        </p>
                      ),
                    )}
                  </div>
                )}
              </>
            )}

            {activeDiscoveryNotice.kind ===
              "new-league" &&
              activeDiscoveryNotice.items.length ===
                0 && (
                <p className="catalogue-update-note">
                  No new catalogue entries were discovered
                  in this refresh. The league collection was
                  still created normally.
                </p>
              )}

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() =>
                  void dismissDiscoveryNotice()
                }
              >
                Got it
              </button>
            </div>
          </section>
        </div>
      )}

      {!activeDiscoveryNotice &&
        catalogueUpdateOpen &&
        catalogueUpdate && (
        <div className="catalogue-update-overlay">
          <section className="catalogue-update-modal">
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                CATALOGUE UPDATED
              </span>
              <h2>{catalogueUpdate.label}</h2>
              <p>
                New Path of Exile 2 catalogue data is available.
                Your existing collection data has been kept.
              </p>
            </div>

            <div className="catalogue-update-stats">
              <div>
                <strong>{catalogueUpdate.newFamilies}</strong>
                <span>new uniques</span>
              </div>
              <div>
                <strong>{catalogueUpdate.newVariants}</strong>
                <span>new variants</span>
              </div>
              <div>
                <strong>{catalogueUpdate.dropDisabled}</strong>
                <span>drop-disabled</span>
              </div>
              <div>
                <strong>{catalogueUpdate.updatedEntries}</strong>
                <span>updated entries</span>
              </div>
            </div>

            <p className="catalogue-update-note">
              New catalogue entries remain Unreviewed until you tell
              PoE 2 Unique Tracker whether you have them.
            </p>

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                onClick={() => acknowledgeCatalogueUpdate(false)}
              >
                Got it
              </button>
              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() => acknowledgeCatalogueUpdate(true)}
              >
                Review New Items
              </button>
            </div>
          </section>
        </div>
      )}

      {pendingImport && (
        <div className="catalogue-update-overlay">
          <section
            className="catalogue-update-modal"
            style={{
              width: "min(760px, calc(100vw - 40px))",
              maxWidth: 760,
            }}
          >
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                IMPORT COLLECTION
              </span>
              <h2>How does this spreadsheet track uniques?</h2>
              <p>
                PoE 2 Unique Tracker detected a likely format, but you choose how the
                spreadsheet should be interpreted before anything is changed.
              </p>
            </div>

            <div className="settings-toggle-list">
              <label className="settings-toggle-row">
                <span>
                  <strong>Full status list</strong>
                  <br />
                  <small>
                    The sheet lists uniques with statuses such as Owned,
                    Wearing, Missing, or Foil.
                  </small>
                </span>
                <input
                  type="radio"
                  name="import-mode"
                  checked={pendingImportMode === "status-list"}
                  onChange={() =>
                    setPendingImportMode("status-list")
                  }
                />
              </label>

              <label className="settings-toggle-row">
                <span>
                  <strong>Missing-only list</strong>
                  {pendingImport.suggestedMode === "missing-only" && (
                    <>
                      {" "}
                      <span className="legacy-badge">DETECTED</span>
                    </>
                  )}
                  <br />
                  <small>
                    Items still written in the sheet are Missing. Older,
                    single-version catalogue uniques absent from the sheet are
                    treated as Owned.
                  </small>
                </span>
                <input
                  type="radio"
                  name="import-mode"
                  checked={pendingImportMode === "missing-only"}
                  onChange={() =>
                    setPendingImportMode("missing-only")
                  }
                />
              </label>


              <button
                type="button"
                className="catalogue-update-secondary"
                style={{
                  justifySelf: "start",
                  marginTop: 2,
                }}
                onClick={() =>
                  setOtherImportModesOpen(
                    (current) => !current,
                  )
                }
              >
                Other
                {pendingImport.suggestedMode ===
                  "color-coded-list" && (
                  <>
                    {" "}
                    <span className="legacy-badge">
                      DETECTED
                    </span>
                  </>
                )}
              </button>

              {otherImportModesOpen && (
                <label className="settings-toggle-row">
                  <span>
                    <strong>Color-coded list</strong>
                    {pendingImport.suggestedMode ===
                      "color-coded-list" && (
                      <>
                        {" "}
                        <span className="legacy-badge">
                          DETECTED
                        </span>
                      </>
                    )}
                    <br />
                    <small>
                      Special format using black/white, light-green text,
                      and green-on-green cells. Unknown ??? entries stay
                      Unreviewed.
                    </small>
                  </span>
                  <input
                    type="radio"
                    name="import-mode"
                    checked={
                      pendingImportMode ===
                      "color-coded-list"
                    }
                    onChange={() => {
                      setPendingImportMode(
                        "color-coded-list",
                      );

                      if (
                        pendingImportProfileId ===
                        STANDARD_PROFILE_ID
                      ) {
                        setPendingImportProfileId(
                          collectionProfiles.find(
                            (profile) =>
                              profile.kind ===
                              "challenge",
                          )?.id ??
                            STANDARD_PROFILE_ID,
                        );
                      }
                    }}
                  />
                </label>
              )}
            </div>

            <div
              style={{
                marginTop: 14,
                display: "grid",
                gap: 6,
              }}
            >
              <label
                className="settings-help"
                htmlFor="spreadsheet-import-destination"
              >
                {pendingImportMode ===
                "color-coded-list"
                  ? "Light-green entries belong to"
                  : "Save imported collection to"}
              </label>
              <select
                id="spreadsheet-import-destination"
                className="collection-profile-select"
                value={pendingImportProfileId}
                disabled={isImporting}
                onChange={(event) =>
                  setPendingImportProfileId(
                    event.target.value,
                  )
                }
              >
                {pendingImportMode ===
                  "color-coded-list" &&
                  pendingImportProfileId ===
                    STANDARD_PROFILE_ID && (
                    <option
                      value={
                        STANDARD_PROFILE_ID
                      }
                      disabled
                    >
                      Choose a challenge league
                    </option>
                  )}

                {collectionProfiles
                  .filter(
                    (profile) =>
                      pendingImportMode !==
                        "color-coded-list" ||
                      profile.kind ===
                        "challenge",
                  )
                  .map((profile) => (
                    <option
                      key={profile.id}
                      value={profile.id}
                    >
                      {profile.name}
                    </option>
                  ))}
              </select>
              <p
                className="catalogue-update-note"
                style={{ margin: 0 }}
              >
                {pendingImportMode ===
                "color-coded-list"
                  ? "Black entries become Missing in every active collection. Light green becomes Owned only in this league. Green-on-green becomes Owned only in Standard. Archived league history is left untouched."
                  : "The spreadsheet and any safe missing-list inference will only change this collection."}
              </p>
            </div>

            {pendingImportMode ===
              "color-coded-list" && (
              <div
                style={{
                  marginTop: 14,
                  border: "1px solid rgba(255,255,255,0.12)",
                  borderRadius: 10,
                  padding: 14,
                }}
              >
                <h3 style={{ marginTop: 0 }}>
                  Color rules
                </h3>

                <p className="settings-help">
                  Black on white = Missing everywhere. Light-green text on
                  white = Owned only in the selected league. Dark-green text
                  on a green background = Owned only in Standard. Rows named
                  ??? are matched by their position between surrounding names
                  and marked Unreviewed rather than Missing or Owned.
                </p>
              </div>
            )}

            {pendingImportMode === "missing-only" && (
              <div
                style={{
                  marginTop: 14,
                  border: "1px solid rgba(255,255,255,0.12)",
                  borderRadius: 10,
                  padding: 14,
                }}
              >
                <h3 style={{ marginTop: 0 }}>
                  New-league safety check
                </h3>

                {pendingImport.latestReleaseLine ? (
                  <>
                    <p className="settings-help">
                      The newest catalogue release is {pendingImport.latestReleaseLine}.
                      Uniques introduced in that release are <strong>not</strong>
                      assumed Owned just because they are absent from this older
                      spreadsheet. PoE 2 Unique Tracker leaves their existing collection
                      state alone instead of inferring Owned.
                    </p>

                    <div className="catalogue-update-stats">
                      <div>
                        <strong>{pendingImport.rows.length}</strong>
                        <span>rows in spreadsheet</span>
                      </div>
                      <div>
                        <strong>
                          {pendingImport.protectedLeagueUniques.length}
                        </strong>
                        <span>new-release uniques protected</span>
                      </div>
                    </div>

                    {pendingImport.protectedLeagueUniques.length > 0 && (
                      <>
                        <button
                          type="button"
                          className="catalogue-update-secondary"
                          onClick={() =>
                            setImportDetailsExpanded(
                              (current) => !current,
                            )
                          }
                        >
                          {importDetailsExpanded
                            ? "Hide New Uniques"
                            : "Show New Uniques"}
                        </button>

                        {importDetailsExpanded && (
                          <div
                            style={{
                              maxHeight: 220,
                              overflowY: "auto",
                              marginTop: 10,
                            }}
                          >
                            {pendingImport.protectedLeagueUniques.map(
                              (item) => (
                                <p
                                  className="catalogue-update-note"
                                  key={item.id}
                                >
                                  <strong>{item.name}</strong>
                                  {item.variantLabel
                                    ? ` \u2014 ${item.variantLabel}`
                                    : ""}
                                  {" \u2022 "}
                                  {item.itemType}
                                </p>
                              ),
                            )}
                          </div>
                        )}
                      </>
                    )}
                  </>
                ) : (
                  <p className="catalogue-update-note">
                    PoE 2 Unique Tracker could not determine the newest release from
                    catalogue metadata. For safety, it will not infer any Owned
                    items from absence until that information is available.
                  </p>
                )}

                <p className="catalogue-update-note">
                  Legacy-only entries and families with multiple exact variants
                  are also never assumed Owned from absence alone.
                </p>
              </div>
            )}

            <p className="catalogue-update-note">
              File: <strong>{pendingImport.fileName}</strong>
            </p>

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                disabled={isImporting}
                onClick={() => {
                  setPendingImport(null);
                  setImportDetailsExpanded(false);
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                className="catalogue-update-primary"
                disabled={
                  isImporting ||
                  (
                    pendingImportMode ===
                      "color-coded-list" &&
                    pendingImportProfileId ===
                      STANDARD_PROFILE_ID
                  )
                }
                onClick={() =>
                  void confirmPendingImport()
                }
              >
                {isImporting
                  ? "Importing..."
                  : pendingImportMode ===
                      "missing-only"
                    ? "Import Missing-Only List"
                    : pendingImportMode ===
                        "color-coded-list"
                      ? "Import Color-Coded List"
                      : "Import Status List"}
              </button>
            </div>
          </section>
        </div>
      )}

      {missingOnlySummary && (
        <div className="catalogue-update-overlay">
          <section
            className="catalogue-update-modal"
            style={{
              width: "min(760px, calc(100vw - 40px))",
              maxWidth: 760,
            }}
          >
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                MISSING-LIST IMPORT COMPLETE
              </span>
              <h2>Collection inferred safely</h2>
              <p>
                PoE 2 Unique Tracker used the missing list to fill {
                  missingOnlySummary.destinationProfileName
                } without pretending that uncertain or newly released uniques
                are already owned.
              </p>
            </div>

            <div className="catalogue-update-stats">
              <div>
                <strong>{missingOnlySummary.explicitMissing}</strong>
                <span>confirmed Missing</span>
              </div>
              <div>
                <strong>{missingOnlySummary.inferredOwned}</strong>
                <span>inferred Owned</span>
              </div>
              <div>
                <strong>
                  {missingOnlySummary.protectedNewLeague.length}
                </strong>
                <span>new-release protected</span>
              </div>
              <div>
                <strong>{missingOnlySummary.protectedAmbiguous}</strong>
                <span>variant/uncertain protected</span>
              </div>
            </div>

            {missingOnlySummary.latestReleaseLine &&
              missingOnlySummary.protectedNewLeague.length > 0 && (
                <>
                  <p className="catalogue-update-note">
                    These uniques are new in release {missingOnlySummary.latestReleaseLine}
                    and were not found in the spreadsheet, so PoE 2 Unique Tracker did
                    <strong> not</strong> mark them Owned.
                  </p>

                  <button
                    type="button"
                    className="catalogue-update-secondary"
                    onClick={() =>
                      setImportDetailsExpanded(
                        (current) => !current,
                      )
                    }
                  >
                    {importDetailsExpanded
                      ? "Hide New Uniques"
                      : "Show New Uniques"}
                  </button>

                  {importDetailsExpanded && (
                    <div
                      style={{
                        maxHeight: 240,
                        overflowY: "auto",
                        marginTop: 10,
                      }}
                    >
                      {missingOnlySummary.protectedNewLeague.map(
                        (item) => (
                          <p
                            className="catalogue-update-note"
                            key={item.id}
                          >
                            <strong>{item.name}</strong>
                            {item.variantLabel
                              ? ` \u2014 ${item.variantLabel}`
                              : ""}
                            {" \u2022 "}
                            {item.itemType}
                          </p>
                        ),
                      )}
                    </div>
                  )}
                </>
              )}

            <div className="catalogue-update-actions">
              {importMatchSummary &&
                (importMatchSummary.ambiguous > 0 ||
                  importMatchSummary.unmatched > 0) && (
                  <button
                    type="button"
                    className="catalogue-update-secondary"
                    onClick={() => {
                      setMissingOnlySummary(null);
                      setImportDetailsExpanded(false);
                      setBatchImportReviewOpen(true);
                    }}
                  >
                    Review Uncertain Imports
                  </button>
                )}

              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() => {
                  setMissingOnlySummary(null);
                  setImportDetailsExpanded(false);
                }}
              >
                Done
              </button>
            </div>
          </section>
        </div>
      )}

      {colorCodedSummary && (
        <div className="catalogue-update-overlay">
          <section
            className="catalogue-update-modal"
            style={{
              width: "min(760px, calc(100vw - 40px))",
              maxWidth: 760,
            }}
          >
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                COLOR-CODED IMPORT COMPLETE
              </span>

              <h2>
                Standard + {colorCodedSummary.leagueProfileName}
              </h2>

              <p>
                The color-coded spreadsheet was applied across the active
                collections. Archived league history was left untouched.
              </p>
            </div>

            <div className="catalogue-update-stats">
              <div>
                <strong>
                  {colorCodedSummary.standardOwned}
                </strong>
                <span>Standard only</span>
              </div>

              <div>
                <strong>
                  {colorCodedSummary.leagueOwned}
                </strong>
                <span>
                  {colorCodedSummary.leagueProfileName} only
                </span>
              </div>

              <div>
                <strong>
                  {colorCodedSummary.missing}
                </strong>
                <span>Missing everywhere</span>
              </div>

              <div>
                <strong>
                  {
                    colorCodedSummary.unknownCandidatesMarkedUnreviewed
                  }
                </strong>
                <span>set Unreviewed</span>
              </div>
            </div>

            {(
              colorCodedSummary.unresolvedPlaceholders >
                0 ||
              colorCodedSummary.unresolvedNamedRows >
                0
            ) && (
              <p className="catalogue-update-note">
                {colorCodedSummary.unresolvedPlaceholders >
                  0
                  ? `${colorCodedSummary.unresolvedPlaceholders} ??? placeholder(s) could not be tied to a catalogue gap and were left unchanged. `
                  : ""}
                {colorCodedSummary.unresolvedNamedRows >
                  0
                  ? `${colorCodedSummary.unresolvedNamedRows} named row(s) did not match cleanly and were left unchanged.`
                  : ""}
              </p>
            )}

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() =>
                  setColorCodedSummary(null)
                }
              >
                Done
              </button>
            </div>
          </section>
        </div>
      )}

      {batchImportReviewOpen && database && (
  <ImportReviewModal
    database={database}
    profiles={collectionProfiles}
    uniques={uniques}
    statusColors={statusColors}
    collectionRules={collectionRules}
    extraTracking={extraTracking}
    initialProfileId={pendingImportProfileId}
    onSaved={async (profileId) => {
      if (profileId === activeProfileId) {
        await loadCollectionData(
          database,
          activeProfileId,
        );
      }
    }}
    onClose={() =>
      setBatchImportReviewOpen(false)
    }
  />
)}

            {closePromptOpen && (
        <div className="catalogue-update-overlay">
          <section className="catalogue-update-modal">
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                CLOSE TRACKER
              </span>

              <h2>What should the tracker do?</h2>

              <p>
                Minimize keeps the window on your taskbar.
                Closing to tray hides the window while keeping
                the in-game lookup hotkey running.
              </p>
            </div>

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                onClick={() =>
                  setClosePromptOpen(false)
                }
              >
                Cancel
              </button>

              <button
                type="button"
                className="catalogue-update-secondary"
                onClick={() =>
                  void minimizeMainWindow()
                }
              >
                Minimize
              </button>

              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() =>
                  void hideMainWindowToTray()
                }
              >
                Close to Tray
              </button>
            </div>
          </section>
        </div>
      )}

      {settingsOpen && (
        <div
          className="settings-overlay"
          onMouseDown={() => setSettingsOpen(false)}
        >
          <section
            className="settings-modal"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="settings-modal-header">
              <div>
                <h2>Settings</h2>
                <p>Customize how PoE 2 Unique Tracker looks and behaves.</p>
              </div>

              <button
                className="settings-close"
                type="button"
                onClick={() => setSettingsOpen(false)}
              >
                {"\u00D7"}
              </button>
            </div>

            <div className="settings-section">
              <h3>Status Colors</h3>

              <div className="status-color-list">
                {STATUS_KEYS.map((status) => {
                  const currentColor = statusColors[status];
                  const isRainbow =
                    status === "foil" && currentColor === "rainbow";

                  return (
                    <div className="status-color-row" key={status}>
                      <span>{STATUS_LABELS[status]}</span>

                      <div className="status-color-controls">
                        {status === "foil" && (
                          <button
                            type="button"
                            className={`rainbow-choice ${
                              isRainbow ? "selected" : ""
                            }`}
                            onClick={() =>
                              changeStatusColor("foil", "rainbow")
                            }
                          >
                            Rainbow
                          </button>
                        )}

                        <input
                          key={`${status}-${currentColor}`}
                          type="color"
                          defaultValue={
                            isRainbow ? "#9b51e0" : currentColor
                          }
                          onBlur={(event) =>
                            changeStatusColor(
                              status,
                              event.currentTarget.value,
                            )
                          }
                          aria-label={`Choose ${STATUS_LABELS[status]} color`}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>

              <button
                type="button"
                className="reset-settings-button"
                onClick={resetStatusColors}
              >
                Reset Colors
              </button>
            </div>

            <div className="settings-section settings-divider-section">
  <h3>In-Game Hotkey</h3>

  <p className="settings-help">
    Choose the shortcut used to check the
    Path of Exile 2 item under your mouse.
  </p>

  <div
    className={`hotkey-test-status ${
      hotkeyReady ? "ready" : ""
    }`}
  >
    <strong>
      Current hotkey
    </strong>

    <span>
      {formatHotkeyForDisplay(
        lookupHotkey,
      )}
    </span>
  </div>

  {hotkeyRecording ? (
    <button
      autoFocus
      type="button"
      className="reset-settings-button"
      onKeyDown={
        captureLookupHotkey
      }
      onBlur={() =>
        setHotkeyRecording(false)
      }
    >
      Press your new shortcut...
    </button>
  ) : (
    <div className="parser-test-actions">
      <button
        type="button"
        className="reset-settings-button"
        disabled={!database}
        onClick={() => {
          setHotkeySettingMessage("");
          setHotkeyRecording(true);
        }}
      >
        Change Hotkey
      </button>

      <button
        type="button"
        className="reset-settings-button"
        disabled={
          !database ||
          lookupHotkey ===
            DEFAULT_POE_LOOKUP_HOTKEY
        }
        onClick={() =>
          void changeLookupHotkey(
            DEFAULT_POE_LOOKUP_HOTKEY,
          )
        }
      >
        Reset to Default
      </button>
    </div>
  )}

  {hotkeyRecording && (
    <p className="settings-help">
      Press the shortcut now.
      Press Escape to cancel.
    </p>
  )}

  {hotkeySettingMessage && (
    <p className="catalogue-check-status">
      {hotkeySettingMessage}
    </p>
  )}
</div>

            <div className="settings-section settings-divider-section">
              <h3>What Counts as Collected?</h3>

              <p className="settings-help">
                Choose which kinds of copies prevent a unique from being
                considered Missing.
              </p>

              <div className="settings-toggle-list">
                <label className="settings-toggle-row">
                  <span>Owned</span>

                  <input
                    type="checkbox"
                    checked={collectionRules.owned}
                    onChange={(event) =>
                      changeCollectionRule("owned", event.target.checked)
                    }
                  />
                </label>

                <label className="settings-toggle-row">
                  <span>Wearing</span>

                  <input
                    type="checkbox"
                    checked={collectionRules.wearing}
                    onChange={(event) =>
                      changeCollectionRule("wearing", event.target.checked)
                    }
                  />
                </label>

                {EXTRA_VARIANTS.map((variant) => (
                  <label className="settings-toggle-row" key={variant}>
                    <span>{STATUS_LABELS[variant]}</span>

                    <input
                      type="checkbox"
                      checked={collectionRules[variant]}
                      onChange={(event) =>
                        changeCollectionRule(
                          variant,
                          event.target.checked,
                        )
                      }
                    />
                  </label>
                ))}
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Extra Variants to Track</h3>

              <p className="settings-help">
                Hidden variants keep their saved collection data and can be
                enabled again later.
              </p>

              <div className="settings-toggle-list">
                {EXTRA_VARIANTS.map((variant) => (
                  <label className="settings-toggle-row" key={variant}>
                    <span>{STATUS_LABELS[variant]}</span>

                    <input
                      type="checkbox"
                      checked={extraTracking[variant]}
                      onChange={(event) =>
                        changeExtraTracking(
                          variant,
                          event.target.checked,
                        )
                      }
                    />
                  </label>
                ))}
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Catalogue Display</h3>

              <p className="settings-help">
                Legacy variants remain saved in your collection even when
                hidden.
              </p>

              <div className="settings-toggle-list">
                <label className="settings-toggle-row">
                  <span>Show Legacy Variants</span>

                  <input
                    type="checkbox"
                    checked={showLegacyVariants}
                    onChange={(event) =>
                      changeShowLegacyVariants(
                        event.target.checked,
                      )
                    }
                  />
                </label>
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Catalogue Updates</h3>

              <p className="settings-help">
                PoE 2 Unique Tracker checks the PoE 2 Wiki catalogue in the background
                at most once every 24 hours. If a new unique appears during a
                league, the tracker adds it as Unreviewed and tells you. A newly
                detected league forces an immediate refresh instead. Your local
                collection still works normally while offline.
              </p>

              <p className="catalogue-check-status">
                {catalogueCheckMessage}
              </p>

              <button
                type="button"
                className="reset-settings-button"
                disabled={catalogueChecking || !database}
                onClick={() => {
                  if (database) {
                    void runCatalogueCheck(database, true);
                  }
                }}
              >
                {catalogueChecking
                  ? "Checking Catalogue..."
                  : "Check for Catalogue Updates"}
              </button>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Imported Collection Matching</h3>

              <p className="settings-help">
                Imported spreadsheets are reference data only. PoE 2 Unique Tracker
                matches their statuses onto the canonical Path of Exile 2
                catalogue instead of creating catalogue entries from imported
                names.
              </p>

              {importMatchSummary ? (
                <div className="import-match-summary">
                  <div>
                    <strong>{importMatchSummary.matched}</strong>
                    <span>matched</span>
                  </div>
                  <div>
                    <strong>{importMatchSummary.ambiguous}</strong>
                    <span>need review</span>
                  </div>
                  <div>
                    <strong>{importMatchSummary.unmatched}</strong>
                    <span>unmatched</span>
                  </div>
                </div>
              ) : (
                <p className="catalogue-check-status">
                  Import matching will run after the canonical catalogue is
                  available.
                </p>
              )}

                            {importMatchSummary &&
                importMatchSummary.skipped > 0 && (
                  <p className="settings-help import-match-note">
                    {importMatchSummary.skipped} imported{" "}
                    {importMatchSummary.skipped === 1
                      ? "row is"
                      : "rows are"}{" "}
                    being ignored by manual choice.
                  </p>
                )}

              {importMatchSummary &&
                (importMatchSummary.ambiguous > 0 ||
                  importMatchSummary.unmatched > 0) && (
                  <>
                    <p className="settings-help import-match-note">
                      {sourceImportMode ===
                      "color-coded-list"
                        ? "Uncertain named rows from the color-coded import were left unchanged. Manual matching is disabled for this format because a row can affect Standard and the selected challenge league differently."
                        : "Nothing was guessed for uncertain rows. Review them manually to connect the spreadsheet entry to the correct canonical unique."}
                    </p>

                    {sourceImportMode !==
                      "color-coded-list" && (
                      <button
                        type="button"
                        className="reset-settings-button"
                        disabled={!database}
                        onClick={
                          openImportReviewScreen
                        }
                      >
                        Review Matches
                      </button>
                    )}
                  </>
                )}
            </div>

            {import.meta.env.DEV && (
              <>
            <div className="settings-section settings-divider-section">
              <h3>Discovery Popup Tests</h3>

              <p className="settings-help">
                Temporary development tools. These only preview the
                discovery popups and do not change or save collection data.
              </p>

              <div className="parser-test-actions">
                <button
                  type="button"
                  className="reset-settings-button"
                  onClick={showDevNewLeagueNotice}
                >
                  Test New League Popup
                </button>

                <button
                  type="button"
                  className="reset-settings-button"
                  onClick={showDevMidLeagueUniqueNotice}
                >
                  Test Mid-League Unique Popup
                </button>
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Item Parser Test</h3>

              <p className="settings-help">
                Temporary development tool: hover a unique in Path of Exile 2,
                press Ctrl+C, paste the copied item text here, and see whether
                PoE 2 Unique Tracker identifies the exact catalogue variant.
              </p>

              <div
                className={`hotkey-test-status ${
                  hotkeyReady ? "ready" : ""
                }`}
              >
                <strong>
                  In-game hotkey:{" "}
{formatHotkeyForDisplay(
  lookupHotkey,
)}
                </strong>
                <span>
                  {hotkeyBusy
                    ? "Looking up hovered item..."
                    : hotkeyMessage}
                </span>
              </div>

              <textarea
                className="parser-test-textarea"
                value={parserTestText}
                onChange={(event) => {
                  setParserTestText(event.target.value);
                  setParserTestResult(null);
                }}
                placeholder={`Item Class: Staves
Rarity: Unique
The Raven's Flock
Perching Staff
--------
...`}
              />

              <div className="parser-test-actions">
                <button
                  type="button"
                  className="reset-settings-button"
                  disabled={parserTesting || !database}
                  onClick={() =>
                    void readClipboardAndIdentify()
                  }
                >
                  {parserTesting
                    ? "Reading Clipboard..."
                    : "Read Clipboard & Identify"}
                </button>

                <button
                  type="button"
                  className="reset-settings-button"
                  disabled={
                    parserTesting ||
                    !database ||
                    !parserTestText.trim()
                  }
                  onClick={() => void runParserTest()}
                >
                  {parserTesting
                    ? "Identifying Item..."
                    : "Identify Pasted Item"}
                </button>
              </div>

              {parserTestResult && (
                <div
                  className={`parser-test-result parser-${parserTestResult.status}`}
                >
                  {parserTestResult.status === "matched" ? (
                    <>
                      <strong>Match found</strong>
                      <span>{parserTestResult.name}</span>
                      <small>
                        {[
                          parserTestResult.baseType,
                          parserTestResult.variantLabel,
                        ]
                          .filter(Boolean)
                          .join(" \u2022 ") || parserTestResult.itemType}
                      </small>
                      <small>
                        Edition:{" "}
                        {parserTestResult.edition === "foulborn"
                          ? "Foulborn"
                          : parserTestResult.edition === "vestigial"
                            ? "Vestigial"
                            : "Normal"}
                      </small>

                      <small>
                        Matched by{" "}
                        {parserTestResult.matchedBy === "special-rule"
                          ? "special variant rule"
                          : parserTestResult.matchedBy === "variant-label"
                            ? "variant evidence"
                            : "name + base type"}
                      </small>

                      {(() => {
                        const localEntry = uniques.find(
                          (unique) =>
                            unique.id === parserTestResult.uniqueId,
                        );

                        if (!localEntry) {
                          return null;
                        }

                        const editionFlag =
                          parserTestResult.edition === "foulborn"
                            ? "foulborn"
                            : parserTestResult.edition === "vestigial"
                              ? "vestigial"
                              : null;

                        if (!localEntry.reviewed) {
                          return (
                            <small>
                              This edition: Unreviewed
                            </small>
                          );
                        }

                        if (editionFlag) {
                          return (
                            <>
                              <small>
                                This edition:{" "}
                                {localEntry.flags.includes(
                                  editionFlag,
                                )
                                  ? "Owned"
                                  : "Missing"}
                              </small>
                              <small>
                                All tracking:{" "}
                                {localEntry.flags.length > 0
                                  ? localEntry.flags.join(", ")
                                  : "none"}
                              </small>
                            </>
                          );
                        }

                        return (
                          <small>
                            Collection:{" "}
                            {localEntry.flags.length > 0
                              ? localEntry.flags.join(", ")
                              : "Missing"}
                          </small>
                        );
                      })()}
                    </>
                  ) : (
                    <>
                      <strong>
                        {parserTestResult.status === "ambiguous"
                          ? "Ambiguous match"
                          : parserTestResult.status === "not-found"
                            ? "Not found"
                            : parserTestResult.status === "not-unique"
                              ? "Not a unique"
                              : "Could not parse item"}
                      </strong>

                      <span>{parserTestResult.message}</span>

                      {"candidates" in parserTestResult &&
                        parserTestResult.candidates.length > 0 && (
                          <small>
                            Candidates:{" "}
                            {parserTestResult.candidates
                              .map((candidate) =>
                                [
                                  candidate.name,
                                  candidate.baseType,
                                  candidate.variantLabel,
                                ]
                                  .filter(Boolean)
                                  .join(" \u2022 "),
                              )
                              .join(" | ")}
                          </small>
                        )}
                    </>
                  )}
                </div>
              )}
            </div>
              </>
            )}
          </section>
        </div>
      )}
    </main>
  );
}

function App() {
  const params = new URLSearchParams(
    window.location.search,
  );

  if (params.get("overlay") === "1") {
    return <OverlayApp />;
  }

  return <MainApp />;
}

export default App;