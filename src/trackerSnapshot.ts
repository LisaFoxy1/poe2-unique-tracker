import { getVersion } from "@tauri-apps/api/app";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readFile, writeFile } from "@tauri-apps/plugin-fs";
import Database from "@tauri-apps/plugin-sql";
import { read, utils, write } from "xlsx";

const SNAPSHOT_FORMAT = "PoE 2 Unique Tracker Snapshot";
const SNAPSHOT_VERSION = 1;
const STANDARD_PROFILE_ID = "standard";

const TRACKING_FLAGS = [
  "owned",
  "wearing",
  "foil",
  "foulborn",
  "vestigial",
] as const;

type TrackingFlag = (typeof TRACKING_FLAGS)[number];

type SqliteTransactionStatement = {
  sql: string;
  params?: string[];
};

type SnapshotProfile = {
  id: string;
  name: string;
  kind: "standard" | "challenge";
  leagueKey: string;
  isArchived: boolean;
  sortOrder: number;
  createdAt: string;
};

type SnapshotCollectionRow = {
  profileId: string;
  uniqueId: string;
  name: string;
  baseType: string;
  itemType: string;
  variant: string;
  reviewed: boolean;
  flags: TrackingFlag[];
};

type ParsedSnapshot = {
  fileName: string;
  path: string;
  exportedAt: string;
  activeProfileId: string;
  profiles: SnapshotProfile[];
  collection: SnapshotCollectionRow[];
};

export type TrackerSnapshotPreview = {
  fileName: string;
  path: string;
  exportedAt: string;
  activeProfileId: string;
  profiles: Array<{
    id: string;
    name: string;
    kind: "standard" | "challenge";
    isArchived: boolean;
  }>;
  collectionRows: number;
};

export type TrackerSnapshotExportSummary = {
  fileName: string;
  path: string;
  profileCount: number;
  collectionRows: number;
};

export type TrackerSnapshotRestoreSummary = {
  fileName: string;
  profileCount: number;
  collectionRows: number;
  restoredTrackingFlags: number;
  remappedUniqueIds: number;
  skippedUnknownUniques: number;
  activeProfileId: string;
};

function fileNameFromPath(path: string) {
  return path.split(/[\\/]/).pop() ?? path;
}

function yesNo(value: boolean) {
  return value ? "Yes" : "No";
}

function parseBoolean(value: unknown) {
  const normalized = String(value ?? "")
    .trim()
    .toLowerCase();

  return (
    normalized === "yes" ||
    normalized === "true" ||
    normalized === "1"
  );
}

function text(value: unknown) {
  return String(value ?? "").trim();
}

function normalizedMatchPart(value: string) {
  return value
    .normalize("NFKC")
    .trim()
    .toLowerCase();
}

function uniqueFallbackKey(
  name: string,
  itemType: string,
  variant: string,
) {
  return [
    normalizedMatchPart(name),
    normalizedMatchPart(itemType),
    normalizedMatchPart(variant),
  ].join("|||");
}

function getInfoValue(
  rows: unknown[][],
  key: string,
) {
  const match = rows.find(
    (row) =>
      text(row[0]).toLowerCase() ===
      key.toLowerCase(),
  );

  return match ? text(match[1]) : "";
}

function getRequiredSheet(
  workbook: ReturnType<typeof read>,
  name: string,
) {
  const sheet = workbook.Sheets[name];

  if (!sheet) {
    throw new Error(
      `This is not a valid tracker snapshot: the "${name}" sheet is missing.`,
    );
  }

  return sheet;
}

async function parseSnapshotFile(
  path: string,
): Promise<ParsedSnapshot> {
  const fileBytes = await readFile(path);

  const workbook = read(fileBytes, {
    type: "array",
    cellFormula: false,
    cellHTML: false,
    cellStyles: false,
  });

  const infoSheet = getRequiredSheet(
    workbook,
    "Tracker Info",
  );

  const profileSheet = getRequiredSheet(
    workbook,
    "Profiles",
  );

  const collectionSheet = getRequiredSheet(
    workbook,
    "Collection",
  );

  const infoRows =
    utils.sheet_to_json<unknown[]>(
      infoSheet,
      {
        header: 1,
        defval: "",
        raw: false,
      },
    );

  const format = getInfoValue(
    infoRows,
    "Format",
  );

  if (format !== SNAPSHOT_FORMAT) {
    throw new Error(
      "That workbook is not a PoE 2 Unique Tracker snapshot.",
    );
  }

  const version = Number(
    getInfoValue(
      infoRows,
      "Snapshot Version",
    ),
  );

  if (version !== SNAPSHOT_VERSION) {
    throw new Error(
      `Snapshot version ${version || "unknown"} is not supported by this version of PoE 2 Unique Tracker.`,
    );
  }

  const exportedAt =
    getInfoValue(
      infoRows,
      "Exported At",
    );

  const requestedActiveProfileId =
    getInfoValue(
      infoRows,
      "Active Profile ID",
    ) || STANDARD_PROFILE_ID;

  const rawProfiles =
    utils.sheet_to_json<
      Record<string, unknown>
    >(
      profileSheet,
      {
        defval: "",
        raw: false,
      },
    );

  const profiles: SnapshotProfile[] = [];
  const seenProfileIds = new Set<string>();

  for (const row of rawProfiles) {
    const id = text(row["Profile ID"]);
    const name = text(row["Name"]);
    const kind = text(row["Kind"]);

    if (!id || !name) {
      continue;
    }

    if (
      kind !== "standard" &&
      kind !== "challenge"
    ) {
      throw new Error(
        `Snapshot profile "${name}" has an invalid profile type.`,
      );
    }

    if (seenProfileIds.has(id)) {
      throw new Error(
        `Snapshot contains duplicate profile ID "${id}".`,
      );
    }

    seenProfileIds.add(id);

    const parsedSortOrder =
      Number(row["Sort Order"]);

    profiles.push({
      id,
      name,
      kind,
      leagueKey: text(
        row["League Key"],
      ),
      isArchived: parseBoolean(
        row["Archived"],
      ),
      sortOrder:
        Number.isFinite(
          parsedSortOrder,
        )
          ? parsedSortOrder
          : 0,
      createdAt:
        text(row["Created At"]) ||
        new Date().toISOString(),
    });
  }

  if (
    !profiles.some(
      (profile) =>
        profile.id ===
          STANDARD_PROFILE_ID &&
        profile.kind === "standard",
    )
  ) {
    throw new Error(
      "The tracker snapshot does not contain a valid Standard collection.",
    );
  }

  const rawCollection =
    utils.sheet_to_json<
      Record<string, unknown>
    >(
      collectionSheet,
      {
        defval: "",
        raw: false,
      },
    );

  const profileIds =
    new Set(
      profiles.map(
        (profile) => profile.id,
      ),
    );

  const collection: SnapshotCollectionRow[] =
    [];

  for (const row of rawCollection) {
    const profileId =
      text(row["Profile ID"]);

    const uniqueId =
      text(row["Unique ID"]);

    if (!profileId || !uniqueId) {
      continue;
    }

    if (!profileIds.has(profileId)) {
      throw new Error(
        `Snapshot collection row references unknown profile "${profileId}".`,
      );
    }

    const flags: TrackingFlag[] = [];

    if (parseBoolean(row["Normal"])) {
      flags.push("owned");
    }

    if (parseBoolean(row["Wearing"])) {
      flags.push("wearing");
    }

    if (parseBoolean(row["Foil"])) {
      flags.push("foil");
    }

    if (
      parseBoolean(row["Foulborn"])
    ) {
      flags.push("foulborn");
    }

    if (
      parseBoolean(row["Vestigial"])
    ) {
      flags.push("vestigial");
    }

    collection.push({
      profileId,
      uniqueId,
      name: text(row["Name"]),
      baseType: text(
        row["Base Type"],
      ),
      itemType: text(
        row["Item Type"],
      ),
      variant: text(row["Variant"]),
      reviewed: parseBoolean(
        row["Reviewed"],
      ),
      flags,
    });
  }

  const activeProfileId =
    profiles.some(
      (profile) =>
        profile.id ===
          requestedActiveProfileId &&
        !profile.isArchived,
    )
      ? requestedActiveProfileId
      : STANDARD_PROFILE_ID;

  return {
    fileName: fileNameFromPath(path),
    path,
    exportedAt,
    activeProfileId,
    profiles,
    collection,
  };
}

async function buildTrackerSnapshotWorkbook(
  db: Database,
) {
  const profiles =
    await db.select<
      Array<{
        id: string;
        name: string;
        kind: string;
        league_key: string | null;
        is_archived: number;
        sort_order: number;
        created_at: string;
      }>
    >(`
      SELECT
        id,
        name,
        kind,
        league_key,
        is_archived,
        sort_order,
        created_at
      FROM collection_profiles
      ORDER BY
        is_archived ASC,
        sort_order ASC,
        name ASC
    `);

  if (profiles.length === 0) {
    throw new Error(
      "There are no tracker collections to export.",
    );
  }

  const catalogue =
    await db.select<
      Array<{
        id: string;
        name: string;
        base_type: string | null;
        item_type: string;
        variant_label: string | null;
      }>
    >(`
      SELECT
        v.id,
        v.name,
        v.base_type,
        v.item_type,
        v.variant_label
      FROM unique_variants v
      WHERE
        v.source IN (
          'poewiki',
          'built-in-special'
        )
        AND NOT (
          v.source = 'poewiki'
          AND EXISTS (
            SELECT 1
            FROM unique_variants special
            WHERE
              special.family_id =
                v.family_id
              AND special.source =
                'built-in-special'
          )
        )
      ORDER BY
        v.item_type ASC,
        v.name ASC,
        COALESCE(
          v.variant_label,
          ''
        ) ASC
    `);

  const reviewRows =
    await db.select<
      Array<{
        profile_id: string;
        unique_id: string;
        reviewed: number;
      }>
    >(`
      SELECT
        profile_id,
        unique_id,
        reviewed
      FROM profile_collection_review
    `);

  const trackingRows =
    await db.select<
      Array<{
        profile_id: string;
        unique_id: string;
        flag: string;
      }>
    >(`
      SELECT
        profile_id,
        unique_id,
        flag
      FROM profile_unique_tracking
    `);

  const activeRows =
    await db.select<
      Array<{ value: string }>
    >(`
      SELECT value
      FROM app_meta
      WHERE
        key =
          'active_collection_profile'
    `);

  const activeProfileId =
    activeRows[0]?.value ??
    STANDARD_PROFILE_ID;

  const reviewedByKey =
    new Map<string, boolean>();

  for (const row of reviewRows) {
    reviewedByKey.set(
      `${row.profile_id}|||${row.unique_id}`,
      row.reviewed === 1,
    );
  }

  const flagsByKey =
    new Map<string, Set<string>>();

  for (const row of trackingRows) {
    const key =
      `${row.profile_id}|||${row.unique_id}`;

    const existing =
      flagsByKey.get(key) ??
      new Set<string>();

    existing.add(row.flag);
    flagsByKey.set(key, existing);
  }

  const appVersion = await getVersion();
  const exportedAt =
    new Date().toISOString();

  const workbook =
    utils.book_new();

  const infoSheet =
    utils.aoa_to_sheet([
      ["Key", "Value"],
      ["Format", SNAPSHOT_FORMAT],
      [
        "Snapshot Version",
        String(
          SNAPSHOT_VERSION,
        ),
      ],
      ["App Version", appVersion],
      ["Exported At", exportedAt],
      [
        "Active Profile ID",
        activeProfileId,
      ],
    ]);

  utils.book_append_sheet(
    workbook,
    infoSheet,
    "Tracker Info",
  );

  const profileSheet =
    utils.json_to_sheet(
      profiles.map(
        (profile) => ({
          "Profile ID":
            profile.id,
          Name: profile.name,
          Kind: profile.kind,
          "League Key":
            profile.league_key ??
            "",
          Archived:
            yesNo(
              profile.is_archived ===
                1,
            ),
          "Sort Order":
            profile.sort_order,
          "Created At":
            profile.created_at,
        }),
      ),
    );

  utils.book_append_sheet(
    workbook,
    profileSheet,
    "Profiles",
  );

  const collectionRows =
    profiles.flatMap(
      (profile) =>
        catalogue.map((item) => {
          const key =
            `${profile.id}|||${item.id}`;

          const flags =
            flagsByKey.get(key) ??
            new Set<string>();

          return {
            "Profile ID":
              profile.id,
            "Profile Name":
              profile.name,
            "Unique ID":
              item.id,
            Name: item.name,
            "Base Type":
              item.base_type ??
              "",
            "Item Type":
              item.item_type,
            Variant:
              item.variant_label ??
              "",
            Reviewed:
              yesNo(
                reviewedByKey.get(
                  key,
                ) ?? false,
              ),
            Normal:
              yesNo(
                flags.has("owned"),
              ),
            Wearing:
              yesNo(
                flags.has(
                  "wearing",
                ),
              ),
            Foil:
              yesNo(
                flags.has("foil"),
              ),
            Foulborn:
              yesNo(
                flags.has(
                  "foulborn",
                ),
              ),
            Vestigial:
              yesNo(
                flags.has(
                  "vestigial",
                ),
              ),
          };
        }),
    );

  const collectionSheet =
    utils.json_to_sheet(
      collectionRows,
    );

  utils.book_append_sheet(
    workbook,
    collectionSheet,
    "Collection",
  );

  return {
    workbook,
    exportedAt,
    profileCount:
      profiles.length,
    collectionRows:
      collectionRows.length,
  };
}

function makeBackupFileName(
  exportedAt: string,
) {
  const stamp =
    exportedAt
      .replace(/\.\d{3}Z$/, "")
      .replace(/[-:T]/g, "")
      .slice(0, 14);

  const year = stamp.slice(0, 4);
  const month = stamp.slice(4, 6);
  const day = stamp.slice(6, 8);
  const hour = stamp.slice(8, 10);
  const minute = stamp.slice(10, 12);
  const second = stamp.slice(12, 14);

  return `PoE2-Tracker-Backup-${year}-${month}-${day}-${hour}${minute}${second}.xlsx`;
}

function workbookToBytes(
  workbook: ReturnType<
    typeof utils.book_new
  >,
) {
  const workbookBytes =
    write(workbook, {
      bookType: "xlsx",
      type: "array",
      compression: true,
    }) as ArrayBuffer;

  return new Uint8Array(
    workbookBytes,
  );
}

export async function exportTrackerSnapshot(
  db: Database,
): Promise<
  TrackerSnapshotExportSummary | null
> {
  const snapshot =
    await buildTrackerSnapshotWorkbook(
      db,
    );

  const day =
    snapshot.exportedAt.slice(
      0,
      10,
    );

  const targetPath =
    await save({
      defaultPath:
        `PoE2-Tracker-Snapshot-${day}.xlsx`,
      filters: [
        {
          name:
            "Excel Workbook",
          extensions: ["xlsx"],
        },
      ],
    });

  if (!targetPath) {
    return null;
  }

  await writeFile(
    targetPath,
    workbookToBytes(
      snapshot.workbook,
    ),
  );

  return {
    fileName:
      fileNameFromPath(
        targetPath,
      ),
    path: targetPath,
    profileCount:
      snapshot.profileCount,
    collectionRows:
      snapshot.collectionRows,
  };
}

export async function createTrackerBackup(
  db: Database,
  directory: string,
): Promise<TrackerSnapshotExportSummary> {
  const snapshot =
    await buildTrackerSnapshotWorkbook(
      db,
    );

  const fileName =
    makeBackupFileName(
      snapshot.exportedAt,
    );

  const bytes =
    workbookToBytes(
      snapshot.workbook,
    );

  const path =
    await invoke<string>(
      "write_tracker_backup",
      {
        directory,
        fileName,
        bytes:
          Array.from(bytes),
      },
    );

  return {
    fileName,
    path,
    profileCount:
      snapshot.profileCount,
    collectionRows:
      snapshot.collectionRows,
  };
}

export async function chooseTrackerSnapshot():
  Promise<TrackerSnapshotPreview | null> {
  const selected =
    await open({
      multiple: false,
      directory: false,
      filters: [
        {
          name:
            "PoE 2 Unique Tracker Snapshot",
          extensions: ["xlsx"],
        },
      ],
    });

  if (
    !selected ||
    Array.isArray(selected)
  ) {
    return null;
  }

  const parsed =
    await parseSnapshotFile(
      selected,
    );

  return {
    fileName:
      parsed.fileName,
    path: parsed.path,
    exportedAt:
      parsed.exportedAt,
    activeProfileId:
      parsed.activeProfileId,
    collectionRows:
      parsed.collection.length,
    profiles:
      parsed.profiles.map(
        (profile) => ({
          id: profile.id,
          name: profile.name,
          kind: profile.kind,
          isArchived:
            profile.isArchived,
        }),
      ),
  };
}

export async function restoreTrackerSnapshot(
  db: Database,
  preview: TrackerSnapshotPreview,
): Promise<TrackerSnapshotRestoreSummary> {
  const parsed =
    await parseSnapshotFile(
      preview.path,
    );

  const currentCatalogue =
    await db.select<
      Array<{
        id: string;
        name: string;
        item_type: string;
        variant_label: string | null;
      }>
    >(`
      SELECT
        v.id,
        v.name,
        v.item_type,
        v.variant_label
      FROM unique_variants v
      WHERE
        v.source IN (
          'poewiki',
          'built-in-special'
        )
        AND NOT (
          v.source = 'poewiki'
          AND EXISTS (
            SELECT 1
            FROM unique_variants special
            WHERE
              special.family_id =
                v.family_id
              AND special.source =
                'built-in-special'
          )
        )
    `);

  const knownIds =
    new Set(
      currentCatalogue.map(
        (item) => item.id,
      ),
    );

  const fallbackIds =
    new Map<string, string[]>();

  for (const item of currentCatalogue) {
    const key =
      uniqueFallbackKey(
        item.name,
        item.item_type,
        item.variant_label ??
          "",
      );

    const existing =
      fallbackIds.get(key) ?? [];

    existing.push(item.id);
    fallbackIds.set(
      key,
      existing,
    );
  }

  let remappedUniqueIds = 0;
  let skippedUnknownUniques = 0;

  const resolvedRows =
    new Map<
      string,
      SnapshotCollectionRow & {
        resolvedUniqueId: string;
      }
    >();

  for (
    const row of
    parsed.collection
  ) {
    let resolvedUniqueId =
      row.uniqueId;

    if (
      !knownIds.has(
        resolvedUniqueId,
      )
    ) {
      const fallback =
        fallbackIds.get(
          uniqueFallbackKey(
            row.name,
            row.itemType,
            row.variant,
          ),
        );

      if (
        fallback?.length === 1
      ) {
        resolvedUniqueId =
          fallback[0];
        remappedUniqueIds += 1;
      } else {
        skippedUnknownUniques +=
          1;
        continue;
      }
    }

    resolvedRows.set(
      `${row.profileId}|||${resolvedUniqueId}`,
      {
        ...row,
        resolvedUniqueId,
      },
    );
  }

  const statements:
    SqliteTransactionStatement[] =
    [
      {
        sql: `
          DELETE FROM profile_unique_tracking
        `,
      },
      {
        sql: `
          DELETE FROM profile_collection_review
        `,
      },
      {
        sql: `
          DELETE FROM league_rollover_changes
        `,
      },
      {
        sql: `
          DELETE FROM league_rollovers
        `,
      },
      {
        sql: `
          DELETE FROM collection_profiles
        `,
      },
    ];

  for (
    const profile of
    parsed.profiles
  ) {
    statements.push({
      sql: `
        INSERT INTO collection_profiles (
          id,
          name,
          kind,
          league_key,
          is_archived,
          sort_order,
          created_at
        )
        VALUES (
          ?,
          ?,
          ?,
          NULLIF(?, ''),
          ?,
          ?,
          ?
        )
      `,
      params: [
        profile.id,
        profile.name,
        profile.kind,
        profile.leagueKey,
        profile.isArchived
          ? "1"
          : "0",
        String(
          profile.sortOrder,
        ),
        profile.createdAt,
      ],
    });
  }

  let restoredTrackingFlags = 0;

  for (
    const row of
    resolvedRows.values()
  ) {
    if (row.reviewed) {
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
          row.profileId,
          row.resolvedUniqueId,
        ],
      });
    }

    for (
      const flag of
      row.flags
    ) {
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
          row.profileId,
          row.resolvedUniqueId,
          flag,
        ],
      });

      restoredTrackingFlags +=
        1;
    }
  }

  statements.push(
    {
      sql: `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'active_collection_profile',
          ?
        )
      `,
      params: [
        parsed.activeProfileId,
      ],
    },
    {
      sql: `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'collection_profiles_migrated_v1',
          'yes'
        )
      `,
    },
    {
      sql: `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'source_file',
          ?
        )
      `,
      params: [
        parsed.fileName,
      ],
    },
    {
      sql: `
        DELETE FROM app_meta
        WHERE key =
          'source_import_mode'
      `,
    },
    {
      sql: `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'pending_tracker_discovery_notices_v1',
          '[]'
        )
      `,
    },
    {
      sql: `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES (
          'pending_new_league_detections_v1',
          '[]'
        )
      `,
    },
  );

  await invoke(
    "execute_sqlite_transaction",
    {
      statements,
      commit: true,
    },
  );

  return {
    fileName:
      parsed.fileName,
    profileCount:
      parsed.profiles.length,
    collectionRows:
      resolvedRows.size,
    restoredTrackingFlags,
    remappedUniqueIds,
    skippedUnknownUniques,
    activeProfileId:
      parsed.activeProfileId,
  };
}
