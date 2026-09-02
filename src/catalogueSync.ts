import Database from "@tauri-apps/plugin-sql";
import { fetch } from "@tauri-apps/plugin-http";

export type CatalogueSyncSummary = {
  revision: string;
  label: string;
  newFamilies: number;
  newVariants: number;
  dropDisabled: number;
  updatedEntries: number;
  checkedAt: number;
  remoteItems: number;
  hasVisibleChanges: boolean;
};

export type ImportReconciliationSummary = {
  matched: number;
  ambiguous: number;
  unmatched: number;
  skipped: number;
  total: number;
};

export type ImportReviewCandidate = {
  id: string;
  name: string;
  baseType: string | null;
  itemType: string;
  variantLabel: string | null;
};

export type ImportReviewItem = {
  importId: string;
  name: string;
  itemType: string;
  importStatus: string;
  status: "ambiguous" | "unmatched";
  matchMethod: string;
  candidates: ImportReviewCandidate[];
};

type CargoTitle = Record<string, string | undefined>;

type CargoResponse = {
  cargoquery?: Array<{
    title?: CargoTitle;
  }>;
};

type WikiCategoryResponse = {
  query?: {
    categorymembers?: Array<{
      pageid?: number;
      ns?: number;
      title?: string;
    }>;
  };
  continue?: {
    cmcontinue?: string;
  };
};

type WikiParsedLinksResponse = {
  parse?: {
    links?: Array<{
      ns?: number;
      title?: string;
      "*"?: string;
    }>;
  };
  error?: {
    code?: string;
    info?: string;
  };
};

type FoilSourceData = {
  allNames: Set<string>;
  voidbornNames: Set<string>;
  valdoNames: Set<string>;
  reliquaryNames: Set<string>;
  reliquaryPagesLoaded: number;
};

export type EditionAvailability =
  | "available"
  | "unavailable"
  | "unknown";

export type CollectibleEdition =
  | "foil"
  | "foulborn"
  | "vestigial";

type RemoteUnique = {
  externalKey: string;
  id: string;
  pageId: string;
  name: string;
  baseType: string | null;
  baseItemId: string | null;
  itemType: string;
  releaseVersion: string | null;
  removalVersion: string | null;
  dropEnabled: number;
  dropRestricted: number;
  isReplica: number;
  isInGame: number;
};

type ExistingCanonicalRow = {
  id: string;
  name: string;
  base_type: string | null;
  item_type: string;
  release_version: string | null;
  drop_enabled: number;
  drop_restricted: number;
  is_replica: number;
  is_legacy_only: number;
  removal_version: string | null;
};

type ImportRow = {
  id: string;
  name: string;
  item_type: string;
  status: string;
};

type VariantRow = {
  id: string;
  family_id: string;
  name: string;
  base_type: string | null;
  item_type: string;
  variant_label: string | null;
  source: string;
};

const POE_WIKI_API = "https://www.poewiki.net/w/api.php";
const PAGE_SIZE = 500;
const CHECK_META_KEY = "last_catalogue_check_at";
const REVISION_META_KEY = "live_catalogue_revision";
const CANONICAL_READY_META_KEY = "canonical_catalogue_ready";

const FOULBORN_CATEGORY =
  "Category:Unique items with Foulborn variants";

const VOIDBORN_FOIL_LIST_PAGE =
  "List of Voidborn foil unique items";

const VALDO_FOIL_LIST_PAGE =
  "List of Valdo's Puzzle Box foil maps";

  const FOIL_RELIQUARY_PAGES = [
  "Ancient Reliquary",
  "Timeworn Reliquary",
  "Vaal Reliquary",
  "Forgotten Reliquary",
  "Visceral Reliquary",
  "Archive Reliquary",
  "Shiny Reliquary",
  "Oubliette Reliquary",
  "Cosmic Reliquary",
  "Decaying Reliquary",
  "Lonely Reliquary",
  "Traumatic Reliquary",
  "Reverent Reliquary",
] as const;

/*
 * Some Reliquary pages mention a unique specifically
 * to say that it is NOT in that Reliquary's foil pool.
 *
 * Those references must not become positive evidence.
 */
const RELIQUARY_LINK_EXCLUSIONS =
  new Set<string>([
    "doryani's prototype",
  ]);

const VESTIGIAL_ITEM_TYPES = new Set([
  "Body Armour",
  "Boots",
  "Gloves",
  "Helmet",
  "Shield",
]);

const COLLECTIBLE_EDITIONS: CollectibleEdition[] = [
  "foil",
  "foulborn",
  "vestigial",
];

const CLASS_TO_TRACKER_TYPE: Record<string, string> = {
  // Flasks
  "Life Flask": "Flask",
  "Life Flasks": "Flask",
  "Mana Flask": "Flask",
  "Mana Flasks": "Flask",
  "Hybrid Flask": "Flask",
  "Hybrid Flasks": "Flask",
  "Utility Flask": "Flask",
  "Utility Flasks": "Flask",

  // Jewellery
  Amulet: "Amulet",
  Amulets: "Amulet",
  Ring: "Ring",
  Rings: "Ring",
  Belt: "Belt",
  Belts: "Belt",

  // Weapons
  Claw: "Claw",
  Claws: "Claw",
  Dagger: "Dagger",
  Daggers: "Dagger",
  "Rune Dagger": "Dagger",
  "Rune Daggers": "Dagger",
  Wand: "Wand",
  Wands: "Wand",
  "One Hand Sword": "Sword",
  "One Hand Swords": "Sword",
  "Thrusting One Hand Sword": "Sword",
  "Thrusting One Hand Swords": "Sword",
  "Two Hand Sword": "Sword",
  "Two Hand Swords": "Sword",
  "One Hand Axe": "Axe",
  "One Hand Axes": "Axe",
  "Two Hand Axe": "Axe",
  "Two Hand Axes": "Axe",
  "One Hand Mace": "Mace",
  "One Hand Maces": "Mace",
  Sceptre: "Mace",
  Sceptres: "Mace",
  "Two Hand Mace": "Mace",
  "Two Hand Maces": "Mace",
  Bow: "Bow",
  Bows: "Bow",
  Staff: "Staff",
  Staves: "Staff",
  Warstaff: "Staff",
  Warstaves: "Staff",
  Quiver: "Quiver",
  Quivers: "Quiver",

  // Armour
  Gloves: "Gloves",
  Boots: "Boots",
  "Body Armour": "Body Armour",
  "Body Armours": "Body Armour",
  Helmet: "Helmet",
  Helmets: "Helmet",
  Shield: "Shield",
  Shields: "Shield",

  // Other collector categories
  Map: "Map",
  Maps: "Map",
  Jewel: "Jewel",
  Jewels: "Jewel",
  "Abyss Jewel": "Jewel",
  "Abyss Jewels": "Jewel",
  "Cluster Jewel": "Jewel",
  "Cluster Jewels": "Jewel",
  Contract: "Contract",
  Contracts: "Contract",
  Tincture: "Tincture",
  Tinctures: "Tincture",
};

function normalize(value: string | null | undefined) {
  return (value ?? "")
    .normalize("NFKC")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function decodeHtmlEntities(value: string) {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) =>
      String.fromCodePoint(parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_, decimal: string) =>
      String.fromCodePoint(parseInt(decimal, 10)),
    )
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");
}

function slug(value: string) {
  return (
    value
      .normalize("NFKD")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "") || "unknown"
  );
}

function parseBool(value: string | undefined, fallback: boolean) {
  if (value == null) {
    return fallback;
  }

  const normalized = value.trim().toLowerCase();

  if (["1", "true", "yes"].includes(normalized)) {
    return true;
  }

  if (["0", "false", "no", ""].includes(normalized)) {
    return false;
  }

  return fallback;
}

function mapItemType(
  itemClass: string | undefined,
  classId: string | undefined,
) {
  const candidates = [classId?.trim(), itemClass?.trim()].filter(
    (value): value is string => Boolean(value),
  );

  for (const candidate of candidates) {
    const mapped = CLASS_TO_TRACKER_TYPE[candidate];

    if (mapped) {
      return mapped;
    }
  }

  // Never silently throw away a Unique because its class is new or unusual.
  // Unknown classes remain visible and can be grouped properly during the
  // final catalogue consistency pass.
  return candidates[1] ?? candidates[0] ?? null;
}

function hashText(text: string) {
  let hash = 0x811c9dc5;

  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }

  return (hash >>> 0).toString(16).padStart(8, "0");
}

function buildCargoUrl(offset: number) {
  const params = new URLSearchParams({
    action: "cargoquery",
    format: "json",
    origin: "*",
    tables: "items",
    fields: [
      "items._pageID=page_id",
      "items.name=name",
      "items.base_item=base_item",
      "items.base_item_id=base_item_id",
      "items.class=item_class",
      "items.class_id=class_id",
      "items.release_version=release_version",
      "items.removal_version=removal_version",
      "items.drop_enabled=drop_enabled",
      "items.is_drop_restricted=drop_restricted",
      "items.is_in_game=is_in_game",
    ].join(","),
    // Do NOT filter is_in_game here. Removed/legacy uniques are part of the
    // collector catalogue too; their metadata tells the UI they are legacy.
    where: 'items.rarity_id="unique"',
    order_by: "items._pageID ASC, items.base_item_id ASC",
    limit: String(PAGE_SIZE),
    offset: String(offset),
  });

  return `${POE_WIKI_API}?${params.toString()}`;
}

function buildCategoryMembersUrl(
  category: string,
  continuation?: string,
) {
  const params = new URLSearchParams({
    action: "query",
    format: "json",
    origin: "*",
    list: "categorymembers",
    cmtitle: category,
    cmnamespace: "0",
    cmtype: "page",
    cmprop: "ids|title",
    cmlimit: String(PAGE_SIZE),
  });

  if (continuation) {
    params.set(
      "cmcontinue",
      continuation,
    );
  }

  return `${POE_WIKI_API}?${params.toString()}`;
}

async function fetchWikiCategoryPageIds(
  category: string,
) {
  const pageIds = new Set<string>();

  let continuation:
    | string
    | undefined;

  do {
    const response = await fetch(
      buildCategoryMembersUrl(
        category,
        continuation,
      ),
      {
        method: "GET",
        headers: {
          Accept: "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `PoE Wiki category request failed (${response.status}).`,
      );
    }

    const payload =
      (await response.json()) as WikiCategoryResponse;

    const members =
      payload.query?.categorymembers ?? [];

    for (const member of members) {
      if (
        member.ns !== 0 ||
        member.pageid == null
      ) {
        continue;
      }

      pageIds.add(
        String(member.pageid),
      );
    }

    continuation =
      payload.continue?.cmcontinue;
  } while (continuation);

  return pageIds;
}

function buildWikiParsedLinksUrl(
  pageTitle: string,
) {
  const params = new URLSearchParams({
    action: "parse",
    format: "json",
    formatversion: "2",
    origin: "*",
    page: pageTitle,
    prop: "links",
    redirects: "1",
  });

  return `${POE_WIKI_API}?${params.toString()}`;
}

async function fetchWikiParsedLinkedNames(
  pageTitle: string,
) {
  const response = await fetch(
    buildWikiParsedLinksUrl(
      pageTitle,
    ),
    {
      method: "GET",
      headers: {
        Accept: "application/json",
      },
    },
  );

  if (!response.ok) {
    throw new Error(
      `PoE Wiki parsed-link request for "${pageTitle}" failed (${response.status}).`,
    );
  }

  const payload =
    (await response.json()) as WikiParsedLinksResponse;

  if (!payload.parse) {
    throw new Error(
      payload.error?.info ??
        `PoE Wiki could not parse "${pageTitle}".`,
    );
  }

  const names =
    new Set<string>();

  for (
    const link of
    payload.parse.links ?? []
  ) {
    if (link.ns !== 0) {
      continue;
    }

    const title =
      link.title ??
      link["*"];

    if (!title) {
      continue;
    }

    names.add(
      normalize(
        decodeHtmlEntities(
          title,
        ),
      ),
    );
  }

  return names;
}

async function fetchWikiParsedLinkedNamesBestEffort(
  pageTitle: string,
) {
  try {
    return await fetchWikiParsedLinkedNames(
      pageTitle,
    );
  } catch (error) {
    console.warn(
      `Could not read Foil source page "${pageTitle}".`,
      error,
    );

    return null;
  }
}

function findUnambiguousFoilMatches(
  remote: RemoteUnique[],
  names: Set<string>,
) {
  const variantsByName =
    new Map<string, RemoteUnique[]>();

  for (const item of remote) {
    const name =
      normalize(item.name);

    const existing =
      variantsByName.get(name) ?? [];

    existing.push(item);

    variantsByName.set(
      name,
      existing,
    );
  }

  const matchedIds =
    new Set<string>();

  for (
    const [name, variants] of
    variantsByName
  ) {
    if (!names.has(name)) {
      continue;
    }

    /*
     * Name-only Wiki evidence cannot tell us which
     * canonical version is meant if multiple tracker
     * variants share the same display name.
     *
     * Leave those uncertain until the variant audit.
     */
    if (variants.length !== 1) {
      continue;
    }

    matchedIds.add(
      variants[0].id,
    );
  }

  return matchedIds;
}

async function fetchKnownFoilSourceData():
  Promise<FoilSourceData> {
  /*
   * Every source here is positive evidence only.
   *
   * Missing from a page does NOT mean that a foil
   * cannot exist.
   */
  const [
    voidbornResult,
    valdoResult,
  ] = await Promise.all([
    fetchWikiParsedLinkedNamesBestEffort(
      VOIDBORN_FOIL_LIST_PAGE,
    ),
    fetchWikiParsedLinkedNamesBestEffort(
      VALDO_FOIL_LIST_PAGE,
    ),
  ]);

  const reliquaryResults =
    await Promise.all(
      FOIL_RELIQUARY_PAGES.map(
        async (pageTitle) => ({
          pageTitle,
          names:
            await fetchWikiParsedLinkedNamesBestEffort(
              pageTitle,
            ),
        }),
      ),
    );

  const reliquaryNames =
    new Set<string>();

  let reliquaryPagesLoaded = 0;

  for (
    const result of
    reliquaryResults
  ) {
    if (!result.names) {
      continue;
    }

    reliquaryPagesLoaded += 1;

    for (const name of result.names) {
      if (
        RELIQUARY_LINK_EXCLUSIONS.has(
          name,
        )
      ) {
        continue;
      }

      reliquaryNames.add(name);
    }
  }

  const voidbornNames =
    voidbornResult ??
    new Set<string>();

  const valdoNames =
    valdoResult ??
    new Set<string>();

  if (
    !voidbornResult &&
    !valdoResult &&
    reliquaryPagesLoaded === 0
  ) {
    throw new Error(
      "Every PoE Wiki Foil source failed to load.",
    );
  }

  return {
    allNames:
      new Set<string>([
        ...voidbornNames,
        ...valdoNames,
        ...reliquaryNames,
      ]),
    voidbornNames,
    valdoNames,
    reliquaryNames,
    reliquaryPagesLoaded,
  };
}

export async function ensureEditionAvailabilitySchema(
  db: Database,
) {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS unique_variant_editions (
      variant_id TEXT NOT NULL,
      edition TEXT NOT NULL,
      availability TEXT NOT NULL,
      source TEXT NOT NULL,
      checked_at INTEGER NOT NULL,
      PRIMARY KEY (
        variant_id,
        edition
      )
    )
  `);
}

async function fetchAllRemoteUniques() {
  const remote: RemoteUnique[] = [];
  const seen = new Set<string>();

  for (let offset = 0; ; offset += PAGE_SIZE) {
    const response = await fetch(buildCargoUrl(offset), {
      method: "GET",
      headers: {
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(
        `PoE Wiki catalogue request failed (${response.status}).`,
      );
    }

    const payload = (await response.json()) as CargoResponse;
    const rows = payload.cargoquery ?? [];

    for (const result of rows) {
      const row = result.title ?? {};

      const name = row.name
        ? decodeHtmlEntities(row.name).trim()
        : undefined;

      const pageId = row.page_id?.trim();
      const itemType = mapItemType(
        row.item_class,
        row.class_id,
      );

      if (!name || !pageId || !itemType) {
        continue;
      }

      const baseType = row.base_item
        ? decodeHtmlEntities(row.base_item).trim() || null
        : null;

      const baseItemId = row.base_item_id?.trim() || null;

      const externalKey = [
        pageId,
        baseItemId ?? baseType ?? itemType,
      ].join("|");

      if (seen.has(externalKey)) {
        continue;
      }

      seen.add(externalKey);

      const id = `poewiki:${pageId}:${slug(
        baseItemId ?? baseType ?? itemType,
      )}`;

      remote.push({
        externalKey,
        id,
        pageId,
        name,
        baseType,
        baseItemId,
        itemType,
        releaseVersion:
          row.release_version?.trim() || null,
        removalVersion:
          row.removal_version?.trim() || null,
        dropEnabled:
          parseBool(row.drop_enabled, true) ? 1 : 0,
        dropRestricted:
          parseBool(row.drop_restricted, false) ? 1 : 0,
        isReplica: name.startsWith("Replica ") ? 1 : 0,
        isInGame:
          parseBool(row.is_in_game, true) ? 1 : 0,
      });
    }

    if (rows.length < PAGE_SIZE) {
      break;
    }
  }

  remote.sort((a, b) => {
    const nameComparison =
      a.name.localeCompare(b.name);

    if (nameComparison !== 0) {
      return nameComparison;
    }

    return (a.baseType ?? "").localeCompare(
      b.baseType ?? "",
    );
  });

  return remote;
}

function metadataChanged(
  existing: ExistingCanonicalRow,
  remote: RemoteUnique,
) {
  return (
    normalize(existing.name) !== normalize(remote.name) ||
    normalize(existing.base_type) !==
      normalize(remote.baseType) ||
    existing.item_type !== remote.itemType ||
    normalize(existing.release_version) !==
      normalize(remote.releaseVersion) ||
    normalize(existing.removal_version) !==
      normalize(remote.removalVersion) ||
    existing.drop_enabled !== remote.dropEnabled ||
    existing.drop_restricted !==
      remote.dropRestricted ||
    existing.is_replica !== remote.isReplica ||
    existing.is_legacy_only !==
      (remote.isInGame ? 0 : 1)
  );
}

function levenshteinDistance(a: string, b: string) {
  const left = normalize(a);
  const right = normalize(b);

  if (left === right) {
    return 0;
  }

  if (left.length === 0) {
    return right.length;
  }

  if (right.length === 0) {
    return left.length;
  }

  let previous = Array.from(
    { length: right.length + 1 },
    (_, index) => index,
  );

  for (
    let leftIndex = 1;
    leftIndex <= left.length;
    leftIndex += 1
  ) {
    const current = [leftIndex];

    for (
      let rightIndex = 1;
      rightIndex <= right.length;
      rightIndex += 1
    ) {
      const insertion =
        current[rightIndex - 1] + 1;
      const deletion =
        previous[rightIndex] + 1;
      const substitution =
        previous[rightIndex - 1] +
        (left[leftIndex - 1] ===
        right[rightIndex - 1]
          ? 0
          : 1);

      current[rightIndex] = Math.min(
        insertion,
        deletion,
        substitution,
      );
    }

    previous = current;
  }

  return previous[right.length];
}

function isSafeTypoMatch(
  importedName: string,
  canonicalName: string,
) {
  const left = normalize(importedName);
  const right = normalize(canonicalName);

  if (!left || !right) {
    return false;
  }

  const longest = Math.max(
    left.length,
    right.length,
  );

  const distance = levenshteinDistance(
    left,
    right,
  );

  if (longest < 8) {
    return distance <= 1;
  }

  if (longest < 20) {
    return distance <= 2;
  }

  return (
    distance <= 3 &&
    distance / longest <= 0.12
  );
}

function isTrackableVariant(
  candidate: VariantRow,
  allVariants: VariantRow[],
) {
  if (candidate.source !== "poewiki") {
    return true;
  }

  // When we have explicit special variants for a family (Ralakesh-style),
  // the generic wiki row remains useful metadata but is not itself a
  // collectible choice.
  return !allVariants.some(
    (other) =>
      other.family_id === candidate.family_id &&
      other.source === "built-in-special",
  );
}

function chooseImportMatch(
  imported: ImportRow,
  allVariants: VariantRow[],
) {
  const trackable = allVariants.filter((candidate) =>
    isTrackableVariant(candidate, allVariants),
  );

  const sameType = trackable.filter(
    (candidate) =>
      candidate.item_type === imported.item_type,
  );

  const exactName = sameType.filter(
    (candidate) =>
      normalize(candidate.name) ===
      normalize(imported.name),
  );

  if (exactName.length === 1) {
    return {
      status: "matched" as const,
      candidate: exactName[0],
      method: "exact-name-type",
      candidates: exactName,
    };
  }

  if (exactName.length > 1) {
    return {
      status: "ambiguous" as const,
      candidate: null,
      method: "exact-name-multiple-variants",
      candidates: exactName,
    };
  }

  const fuzzy = sameType
    .filter((candidate) =>
      isSafeTypoMatch(
        imported.name,
        candidate.name,
      ),
    )
    .map((candidate) => ({
      candidate,
      distance: levenshteinDistance(
        imported.name,
        candidate.name,
      ),
    }))
    .sort(
      (a, b) => a.distance - b.distance,
    );

  if (fuzzy.length > 0) {
    const bestDistance = fuzzy[0].distance;
    const best = fuzzy.filter(
      (entry) =>
        entry.distance === bestDistance,
    );

    if (best.length === 1) {
      // A fuzzy match must still point to a family with one trackable
      // collectible variant. Otherwise a typo must not make us guess a
      // Ralakesh-style variant.
      const familyCandidates = trackable.filter(
        (candidate) =>
          candidate.family_id ===
          best[0].candidate.family_id,
      );

      if (familyCandidates.length === 1) {
        return {
          status: "matched" as const,
          candidate: best[0].candidate,
          method: "safe-typo-name-type",
          candidates: [best[0].candidate],
        };
      }

      return {
        status: "ambiguous" as const,
        candidate: null,
        method: "safe-typo-multiple-variants",
        candidates: familyCandidates,
      };
    }
  }

  return {
    status: "unmatched" as const,
    candidate: null,
    method: "no-safe-match",
    candidates: [] as VariantRow[],
  };
}

async function ensureImportReconciliationSchema(
  db: Database,
) {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS import_reconciliation (
      import_id TEXT PRIMARY KEY NOT NULL,
      canonical_variant_id TEXT,
      status TEXT NOT NULL,
      match_method TEXT NOT NULL,
      candidates_json TEXT NOT NULL DEFAULT '[]',
      last_attempt_at INTEGER NOT NULL
    )
  `);

  await db.execute(`
    CREATE TABLE IF NOT EXISTS import_manual_resolutions (
      import_id TEXT PRIMARY KEY NOT NULL,
      canonical_variant_id TEXT,
      action TEXT NOT NULL,
      resolved_at INTEGER NOT NULL
    )
  `);
}

async function copyImportedStateToCanonical(
  db: Database,
  importId: string,
  canonicalVariantId: string,
) {
  const existingFlags =
    await db.select<
      { flag: string }[]
    >(
      `
        SELECT flag
        FROM unique_tracking
        WHERE unique_id = ?
      `,
      [importId],
    );

  for (const { flag } of existingFlags) {
    await db.execute(
      `
        INSERT OR IGNORE INTO unique_tracking (
          unique_id,
          flag
        )
        VALUES (?, ?)
      `,
      [
        canonicalVariantId,
        flag,
      ],
    );
  }

  await db.execute(
    `
      INSERT OR REPLACE INTO collection_review (
        unique_id,
        reviewed
      )
      VALUES (?, 1)
    `,
    [canonicalVariantId],
  );
}

export async function loadImportReviewItems(
  db: Database,
): Promise<ImportReviewItem[]> {
  await ensureImportReconciliationSchema(
    db,
  );

  const rows =
    await db.select<
      {
        import_id: string;
        name: string;
        item_type: string;
        import_status: string;
        reconciliation_status: string;
        match_method: string;
        candidates_json: string;
      }[]
    >(`
      SELECT
        reconciliation.import_id,
        imported.name,
        imported.item_type,
        imported.status AS import_status,
        reconciliation.status AS reconciliation_status,
        reconciliation.match_method,
        reconciliation.candidates_json
      FROM import_reconciliation reconciliation
      JOIN imported_collection imported
        ON imported.id = reconciliation.import_id
      WHERE reconciliation.status IN (
        'ambiguous',
        'unmatched'
      )
      ORDER BY
        imported.name COLLATE NOCASE ASC,
        imported.item_type COLLATE NOCASE ASC
    `);

  return rows.flatMap((row) => {
    if (
      row.reconciliation_status !== "ambiguous" &&
      row.reconciliation_status !== "unmatched"
    ) {
      return [];
    }

    let candidates:
      ImportReviewCandidate[] = [];

    try {
      const parsed =
        JSON.parse(
          row.candidates_json,
        ) as unknown;

      if (Array.isArray(parsed)) {
        candidates =
          parsed.flatMap(
            (candidate) => {
              if (
                !candidate ||
                typeof candidate !== "object"
              ) {
                return [];
              }

              const value =
                candidate as {
                  id?: unknown;
                  name?: unknown;
                  baseType?: unknown;
                  itemType?: unknown;
                  variantLabel?: unknown;
                };

              if (
                typeof value.id !== "string" ||
                typeof value.name !== "string" ||
                typeof value.itemType !== "string"
              ) {
                return [];
              }

              return [
                {
                  id: value.id,
                  name: value.name,
                  baseType:
                    typeof value.baseType === "string"
                      ? value.baseType
                      : null,
                  itemType:
                    value.itemType,
                  variantLabel:
                    typeof value.variantLabel === "string"
                      ? value.variantLabel
                      : null,
                },
              ];
            },
          );
      }
    } catch {
      candidates = [];
    }

    return [
      {
        importId: row.import_id,
        name: row.name,
        itemType: row.item_type,
        importStatus:
          row.import_status,
        status:
          row.reconciliation_status,
        matchMethod:
          row.match_method,
        candidates,
      },
    ];
  });
}

export async function saveManualImportResolution(
  db: Database,
  importId: string,
  canonicalVariantId:
    | string
    | null,
) {
  await ensureImportReconciliationSchema(
    db,
  );

  if (!canonicalVariantId) {
    await db.execute(
      `
        INSERT OR REPLACE INTO import_manual_resolutions (
          import_id,
          canonical_variant_id,
          action,
          resolved_at
        )
        VALUES (?, NULL, 'skipped', ?)
      `,
      [
        importId,
        Date.now(),
      ],
    );

    return;
  }

  const canonicalRows =
    await db.select<
      { id: string }[]
    >(
      `
        SELECT id
        FROM unique_variants
        WHERE
          id = ?
          AND source IN (
            'poewiki',
            'built-in-special'
          )
      `,
      [canonicalVariantId],
    );

  if (canonicalRows.length === 0) {
    throw new Error(
      "That catalogue variant no longer exists.",
    );
  }

  await db.execute(
    `
      INSERT OR REPLACE INTO import_manual_resolutions (
        import_id,
        canonical_variant_id,
        action,
        resolved_at
      )
      VALUES (?, ?, 'matched', ?)
    `,
    [
      importId,
      canonicalVariantId,
      Date.now(),
    ],
  );

  await copyImportedStateToCanonical(
    db,
    importId,
    canonicalVariantId,
  );

  await db.execute(
    `
      INSERT OR REPLACE INTO profile_collection_review (
        profile_id,
        unique_id,
        reviewed
      )
      VALUES ('standard', ?, 1)
    `,
    [canonicalVariantId],
  );

  const importedFlags =
    await db.select<
      { flag: string }[]
    >(
      `
        SELECT flag
        FROM unique_tracking
        WHERE unique_id = ?
      `,
      [importId],
    );

  for (const { flag } of importedFlags) {
    await db.execute(
      `
        INSERT OR IGNORE INTO profile_unique_tracking (
          profile_id,
          unique_id,
          flag
        )
        VALUES (
          'standard',
          ?,
          ?
        )
      `,
      [
        canonicalVariantId,
        flag,
      ],
    );
  }
}

export async function shouldCheckCatalogueNow(
  db: Database,
  maxAgeMs = 24 * 60 * 60 * 1000,
) {
  const rows = await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = '${CHECK_META_KEY}'
  `);

  if (rows.length === 0) {
    return true;
  }

  const lastChecked = Number(
    rows[0].value,
  );

  if (!Number.isFinite(lastChecked)) {
    return true;
  }

  return (
    Date.now() - lastChecked >= maxAgeMs
  );
}

export async function isCanonicalCatalogueReady(
  db: Database,
) {
  const rows = await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = '${CANONICAL_READY_META_KEY}'
  `);

  return rows[0]?.value === "yes";
}

export async function syncCatalogueFromPoeWiki(
  db: Database,
): Promise<CatalogueSyncSummary> {
  // Fetch the COMPLETE remote catalogue before any write. A network failure
  // therefore leaves the local catalogue untouched.
  const remote =
  await fetchAllRemoteUniques();

let foulbornPageIds:
  | Set<string>
  | null = null;

try {
  foulbornPageIds =
    await fetchWikiCategoryPageIds(
      FOULBORN_CATEGORY,
    );
} catch (error) {
  console.warn(
    "Could not update Foulborn availability. Existing edition data will be kept:",
    error,
  );
}

let foilSourceData:
  | FoilSourceData
  | null = null;

try {
  foilSourceData =
    await fetchKnownFoilSourceData();
} catch (error) {
  console.warn(
    "Could not update Foil availability. Existing edition data will be kept:",
    error,
  );
}

if (remote.length < 800) {
    throw new Error(
      `PoE Wiki returned only ${remote.length} usable unique entries. ` +
        "That is far below the expected catalogue size, so the local " +
        "catalogue was not changed.",
    );
  }

  await db.execute(`
    CREATE TABLE IF NOT EXISTS catalogue_external_links (
      source TEXT NOT NULL,
      external_key TEXT NOT NULL,
      unique_id TEXT NOT NULL,
      PRIMARY KEY (source, external_key)
    )
  `);

  await ensureEditionAvailabilitySchema(
  db,
);

  const revisionPayload = remote.map(
    (item) => [
      item.externalKey,
      item.name,
      item.baseType,
      item.itemType,
      item.releaseVersion,
      item.removalVersion,
      item.dropEnabled,
      item.dropRestricted,
      item.isReplica,
      item.isInGame,
    ],
  );

  const revision = `poewiki-${
    remote.length
  }-${hashText(
    JSON.stringify(revisionPayload),
  )}`;

  const previousRevisionRows =
    await db.select<{ value: string }[]>(`
      SELECT value
      FROM app_meta
      WHERE key = '${REVISION_META_KEY}'
    `);

  const previousRevision =
    previousRevisionRows[0]?.value ?? null;

  const existingRows =
    await db.select<
      ExistingCanonicalRow[]
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
        is_legacy_only,
        removal_version
      FROM unique_catalogue
      WHERE source = 'poewiki'
    `);

  const existingById = new Map(
    existingRows.map((row) => [
      row.id,
      row,
    ]),
  );

  const existingFamilyNames = new Set(
    existingRows.map((row) =>
      normalize(row.name),
    ),
  );

  let newFamilies = 0;
  let newVariants = 0;
  let dropDisabled = 0;
  let updatedEntries = 0;

  for (const item of remote) {
    const existing =
      existingById.get(item.id) ?? null;

    if (!existing) {
      if (
        existingFamilyNames.has(
          normalize(item.name),
        )
      ) {
        newVariants += 1;
      } else {
        newFamilies += 1;
        existingFamilyNames.add(
          normalize(item.name),
        );
      }
    } else {
      if (
        existing.drop_enabled === 1 &&
        item.dropEnabled === 0
      ) {
        dropDisabled += 1;
      }

      if (
        metadataChanged(existing, item)
      ) {
        updatedEntries += 1;
      }
    }

    await db.execute(
      `
        INSERT OR REPLACE INTO unique_catalogue (
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
        )
        VALUES (
          ?, ?, ?, ?, ?, ?, ?, ?,
          COALESCE(
            (
              SELECT has_legacy_variants
              FROM unique_catalogue
              WHERE id = ?
            ),
            0
          ),
          ?, ?, 'poewiki'
        )
      `,
      [
        item.id,
        item.name,
        item.baseType,
        item.itemType,
        item.releaseVersion,
        item.dropEnabled,
        item.dropRestricted,
        item.isReplica,
        item.id,
        item.isInGame ? 0 : 1,
        item.removalVersion,
      ],
    );
  }

  // The query includes current + removed uniques, so stale poewiki rows really
  // are stale and can be removed after the complete fetch succeeded.
  const remoteIds = new Set(
    remote.map((item) => item.id),
  );

  for (const existing of existingRows) {
    if (!remoteIds.has(existing.id)) {
      await db.execute(
        `
          DELETE FROM unique_catalogue
          WHERE id = ? AND source = 'poewiki'
        `,
        [existing.id],
      );
    }
  }

  const editionCheckedAt =
  Date.now();

/*
 * Every canonical Wiki variant gets a row for every
 * collectible edition.
 *
 * INSERT OR IGNORE is deliberate: if we already know
 * something more specific, merely creating the schema
 * must not overwrite it.
 */
for (const item of remote) {
  for (const edition of COLLECTIBLE_EDITIONS) {
    await db.execute(
      `
        INSERT OR IGNORE INTO unique_variant_editions (
          variant_id,
          edition,
          availability,
          source,
          checked_at
        )
        VALUES (?, ?, 'unknown', 'default', ?)
      `,
      [
        item.id,
        edition,
        editionCheckedAt,
      ],
    );
  }
}

/*
 * The Wiki maintains an explicit Foulborn category.
 * When that complete category fetch succeeds, it is
 * authoritative enough for our automated Foulborn
 * availability layer.
 *
 * Manual overrides are protected so proofreading
 * corrections can take precedence later.
 */
if (foulbornPageIds) {
  let foulbornAvailableCount = 0;

  for (const item of remote) {
    const availability:
      EditionAvailability =
      foulbornPageIds.has(
        item.pageId,
      )
        ? "available"
        : "unavailable";

    if (
      availability === "available"
    ) {
      foulbornAvailableCount += 1;
    }

    await db.execute(
      `
        INSERT INTO unique_variant_editions (
          variant_id,
          edition,
          availability,
          source,
          checked_at
        )
        VALUES (
          ?,
          'foulborn',
          ?,
          'poewiki-foulborn-category',
          ?
        )
        ON CONFLICT (
          variant_id,
          edition
        )
        DO UPDATE SET
          availability =
            excluded.availability,
          source =
            excluded.source,
          checked_at =
            excluded.checked_at
        WHERE
          unique_variant_editions.source !=
            'manual'
      `,
      [
        item.id,
        availability,
        editionCheckedAt,
      ],
    );
  }

  console.log(
    `Foulborn availability updated: ${foulbornAvailableCount} catalogue variants marked available from ${foulbornPageIds.size} Wiki pages.`,
  );
}

/*
 * Foil colour and acquisition source do NOT create
 * separate tracker editions.
 *
 * Presence on an explicit Voidborn / Valdo foil list
 * is enough to CONFIRM that a foil exists.
 *
 * Absence from those lists is not enough to prove that
 * a foil does not exist, so unmatched uniques remain
 * unknown rather than being hidden.
 */

/*
 * Older versions of PoE Collector treated "not found"
 * as unavailable. Undo those old guesses first.
 */
await db.execute(
  `
    UPDATE unique_variant_editions
    SET
      availability = 'unknown',
      source = 'poewiki-foil-unconfirmed',
      checked_at = ?
    WHERE
      edition = 'foil'
      AND source IN (
        'poewiki-foil-sources',
        'poewiki-foil-confirmed',
        'poewiki-foil-unconfirmed'
      )
  `,
  [editionCheckedAt],
);

if (foilSourceData) {
  const voidbornMatchedIds =
  findUnambiguousFoilMatches(
    remote,
    foilSourceData.voidbornNames,
  );

const valdoMatchedIds =
  findUnambiguousFoilMatches(
    remote,
    foilSourceData.valdoNames,
  );

const reliquaryMatchedIds =
  findUnambiguousFoilMatches(
    remote,
    foilSourceData.reliquaryNames,
  );

const foilMatchedIds =
  findUnambiguousFoilMatches(
    remote,
    foilSourceData.allNames,
  );

  for (const item of remote) {
    if (!foilMatchedIds.has(item.id)) {
      continue;
    }

    await db.execute(
      `
        INSERT INTO unique_variant_editions (
          variant_id,
          edition,
          availability,
          source,
          checked_at
        )
        VALUES (
          ?,
          'foil',
          'available',
          'poewiki-foil-confirmed',
          ?
        )
        ON CONFLICT (
          variant_id,
          edition
        )
        DO UPDATE SET
          availability = 'available',
          source = 'poewiki-foil-confirmed',
          checked_at =
            excluded.checked_at
        WHERE
          unique_variant_editions.source !=
            'manual'
      `,
      [
        item.id,
        editionCheckedAt,
      ],
    );
  }

  console.log(
  `Foil availability updated: ${foilMatchedIds.size} catalogue variants confirmed. ` +
    `Voidborn matches: ${voidbornMatchedIds.size}. ` +
    `Valdo matches: ${valdoMatchedIds.size}. ` +
    `Reliquary matches: ${reliquaryMatchedIds.size}. ` +
    `Reliquary pages loaded: ${foilSourceData.reliquaryPagesLoaded}/${FOIL_RELIQUARY_PAGES.length}. ` +
    `Unmatched or ambiguous variants remain unknown.`,
);
}

/*
 * Vestigial uniques are restricted by item class.
 *
 * The five supported tracker classes are:
 * Body Armour, Boots, Gloves, Helmet, and Shield.
 *
 * This is considered conclusive enough to hide the
 * Vestigial button from every other item class.
 */
for (const item of remote) {
  const availability:
    EditionAvailability =
    VESTIGIAL_ITEM_TYPES.has(
      item.itemType,
    )
      ? "available"
      : "unavailable";

  await db.execute(
    `
      INSERT INTO unique_variant_editions (
        variant_id,
        edition,
        availability,
        source,
        checked_at
      )
      VALUES (
        ?,
        'vestigial',
        ?,
        'built-in-vestigial-class-rule',
        ?
      )
      ON CONFLICT (
        variant_id,
        edition
      )
      DO UPDATE SET
        availability =
          excluded.availability,
        source =
          excluded.source,
        checked_at =
          excluded.checked_at
      WHERE
        unique_variant_editions.source !=
          'manual'
    `,
    [
      item.id,
      availability,
      editionCheckedAt,
    ],
  );
}

/*
 * Built-in special catalogue variants, such as the
 * explicit Ralakesh variants, do not appear in the
 * remote Wiki array above. Apply the same item-class
 * rule to those too.
 */
const builtInSpecialVariants =
  await db.select<
    {
      id: string;
      item_type: string;
    }[]
  >(`
    SELECT
      id,
      item_type
    FROM unique_variants
    WHERE source = 'built-in-special'
  `);

for (
  const variant of
  builtInSpecialVariants
) {
  const availability:
    EditionAvailability =
    VESTIGIAL_ITEM_TYPES.has(
      variant.item_type,
    )
      ? "available"
      : "unavailable";

  await db.execute(
    `
      INSERT INTO unique_variant_editions (
        variant_id,
        edition,
        availability,
        source,
        checked_at
      )
      VALUES (
        ?,
        'vestigial',
        ?,
        'built-in-vestigial-class-rule',
        ?
      )
      ON CONFLICT (
        variant_id,
        edition
      )
      DO UPDATE SET
        availability =
          excluded.availability,
        source =
          excluded.source,
        checked_at =
          excluded.checked_at
      WHERE
        unique_variant_editions.source !=
          'manual'
    `,
    [
      variant.id,
      availability,
      editionCheckedAt,
    ],
  );
}

console.log(
  "Vestigial availability updated using item-class rules.",
);

await db.execute(`
  DELETE FROM unique_variant_editions
  WHERE
    variant_id LIKE 'poewiki:%'
    AND variant_id NOT IN (
      SELECT id
      FROM unique_catalogue
      WHERE source = 'poewiki'
    )
`);

  await db.execute(`
    DELETE FROM catalogue_external_links
    WHERE source = 'poewiki'
  `);

  for (const item of remote) {
    await db.execute(
      `
        INSERT INTO catalogue_external_links (
          source,
          external_key,
          unique_id
        )
        VALUES ('poewiki', ?, ?)
      `,
      [item.externalKey, item.id],
    );
  }

  const checkedAt = Date.now();

  await db.execute(
    `
      INSERT OR REPLACE INTO app_meta (
        key,
        value
      )
      VALUES ('${CHECK_META_KEY}', ?)
    `,
    [String(checkedAt)],
  );

  await db.execute(
    `
      INSERT OR REPLACE INTO app_meta (
        key,
        value
      )
      VALUES ('${REVISION_META_KEY}', ?)
    `,
    [revision],
  );

  await db.execute(`
    INSERT OR REPLACE INTO app_meta (
      key,
      value
    )
    VALUES (
      '${CANONICAL_READY_META_KEY}',
      'yes'
    )
  `);

  const hasVisibleChanges =
    previousRevision !== null &&
    previousRevision !== revision &&
    newFamilies +
      newVariants +
      dropDisabled +
      updatedEntries >
      0;

  const firstCanonicalSync =
    previousRevision === null &&
    newFamilies + newVariants > 0;

  return {
    revision,
    label:
      "Canonical PoE Wiki catalogue",
    newFamilies,
    newVariants,
    dropDisabled,
    updatedEntries,
    checkedAt,
    remoteItems: remote.length,
    hasVisibleChanges:
      hasVisibleChanges ||
      firstCanonicalSync,
  };
}

export async function reconcileImportedCollection(
  db: Database,
): Promise<ImportReconciliationSummary> {
  await ensureImportReconciliationSchema(
  db,
);

  const canonicalReady =
    await isCanonicalCatalogueReady(db);

  if (!canonicalReady) {
    return {
  matched: 0,
  ambiguous: 0,
  unmatched: 0,
  skipped: 0,
  total: 0,
};
  }

  const imports =
    await db.select<ImportRow[]>(`
      SELECT id, name, item_type, status
      FROM imported_collection
    `);

  const allVariants =
    await db.select<VariantRow[]>(`
      SELECT
        id,
        family_id,
        name,
        base_type,
        item_type,
        variant_label,
        source
      FROM unique_variants
      WHERE source != 'imported'
    `);

      const manualRows =
    await db.select<
      {
        import_id: string;
        canonical_variant_id:
          | string
          | null;
        action: string;
      }[]
    >(`
      SELECT
        import_id,
        canonical_variant_id,
        action
      FROM import_manual_resolutions
    `);

  const manualByImport =
    new Map(
      manualRows.map(
        (row) => [
          row.import_id,
          row,
        ] as const,
      ),
    );

  const variantById =
    new Map(
      allVariants.map(
        (variant) => [
          variant.id,
          variant,
        ] as const,
      ),
    );

    let matched = 0;
    let ambiguous = 0;
    let unmatched = 0;
    let skipped = 0;

  const now = Date.now();

    for (const imported of imports) {
    const manual =
      manualByImport.get(
        imported.id,
      );

    if (
      manual?.action === "skipped"
    ) {
      skipped += 1;

      await db.execute(
        `
          INSERT OR REPLACE INTO import_reconciliation (
            import_id,
            canonical_variant_id,
            status,
            match_method,
            candidates_json,
            last_attempt_at
          )
          VALUES (
            ?,
            NULL,
            'skipped',
            'manual-skip',
            '[]',
            ?
          )
        `,
        [
          imported.id,
          now,
        ],
      );

      continue;
    }

    if (
      manual?.action === "matched" &&
      manual.canonical_variant_id
    ) {
      const manualCandidate =
        variantById.get(
          manual.canonical_variant_id,
        );

      if (
        manualCandidate &&
        isTrackableVariant(
          manualCandidate,
          allVariants,
        )
      ) {
        matched += 1;

        await copyImportedStateToCanonical(
          db,
          imported.id,
          manualCandidate.id,
        );

        await db.execute(
          `
            INSERT OR REPLACE INTO import_reconciliation (
              import_id,
              canonical_variant_id,
              status,
              match_method,
              candidates_json,
              last_attempt_at
            )
            VALUES (
              ?,
              ?,
              'matched',
              'manual-resolution',
              '[]',
              ?
            )
          `,
          [
            imported.id,
            manualCandidate.id,
            now,
          ],
        );

        continue;
      }

      await db.execute(
        `
          DELETE FROM import_manual_resolutions
          WHERE import_id = ?
        `,
        [imported.id],
      );
    }

    const result = chooseImportMatch(
      imported,
      allVariants,
    );

    const candidateJson = JSON.stringify(
      result.candidates.map(
        (candidate) => ({
          id: candidate.id,
          name: candidate.name,
          baseType:
            candidate.base_type,
          itemType:
            candidate.item_type,
          variantLabel:
            candidate.variant_label,
        }),
      ),
    );

    if (
  result.status === "matched" &&
  result.candidate
) {
  matched += 1;

  await copyImportedStateToCanonical(
    db,
    imported.id,
    result.candidate.id,
  );

  await db.execute(
    `
      INSERT OR REPLACE INTO import_reconciliation (
        import_id,
        canonical_variant_id,
        status,
        match_method,
        candidates_json,
        last_attempt_at
      )
      VALUES (?, ?, 'matched', ?, ?, ?)
    `,
    [
      imported.id,
      result.candidate.id,
      result.method,
      candidateJson,
      now,
    ],
  );
} else if (
  result.status === "ambiguous"
    ) {
      ambiguous += 1;

      await db.execute(
        `
          INSERT OR REPLACE INTO import_reconciliation (
            import_id,
            canonical_variant_id,
            status,
            match_method,
            candidates_json,
            last_attempt_at
          )
          VALUES (?, NULL, 'ambiguous', ?, ?, ?)
        `,
        [
          imported.id,
          result.method,
          candidateJson,
          now,
        ],
      );
    } else {
      unmatched += 1;

      await db.execute(
        `
          INSERT OR REPLACE INTO import_reconciliation (
            import_id,
            canonical_variant_id,
            status,
            match_method,
            candidates_json,
            last_attempt_at
          )
          VALUES (?, NULL, 'unmatched', ?, '[]', ?)
        `,
        [
          imported.id,
          result.method,
          now,
        ],
      );
    }
  }

  // Imported rows are reference/history only. Once canonical reconciliation
  // has run, they must never remain visible as catalogue items.
  await db.execute(`
    DELETE FROM unique_variants
    WHERE source NOT IN ('poewiki', 'built-in-special')
  `);

  await db.execute(`
    DELETE FROM unique_catalogue
    WHERE source != 'poewiki'
  `);

  await db.execute(`
    DELETE FROM unique_families
    WHERE id NOT IN (
      SELECT DISTINCT family_id
      FROM unique_variants
    )
  `);

    return {
    matched,
    ambiguous,
    unmatched,
    skipped,
    total: imports.length,
  };
}