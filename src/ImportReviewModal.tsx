import { useEffect, useState, type CSSProperties } from "react";
import Database from "@tauri-apps/plugin-sql";
import { invoke } from "@tauri-apps/api/core";
import { reconcileImportedCollection } from "./catalogueSync";

type Flag = "owned" | "wearing" | "foil" | "foulborn" | "vestigial";
type Extra = "foil" | "foulborn" | "vestigial";
type Availability = "available" | "unavailable" | "unknown";
type Profile = { id: string; name: string };
type Unique = {
  id: string;
  name: string;
  baseType: string | null;
  itemType: string;
  variantLabel: string | null;
  editionAvailability: Record<Extra, Availability>;
};
type Candidate = Pick<Unique, "id" | "name" | "baseType" | "itemType" | "variantLabel">;
type ReviewItem = {
  importId: string;
  name: string;
  itemType: string;
  importStatus: string;
  status: "ambiguous" | "unmatched";
  candidates: Candidate[];
};
type ProfileStatus = { reviewed: boolean; flags: Flag[] };
type Decision = "pending" | "reviewed" | "ignored";
type DraftRow = { decision: Decision; changes: Record<string, Flag[]> };
type Draft = Record<string, DraftRow>;
type SqliteTransactionStatement = {
  sql: string;
  params?: string[];
};

type Props = {
  database: Database;
  profiles: Profile[];
  uniques: Unique[];
  statusColors: Record<"missing" | Flag, string>;
  collectionRules: Record<Flag, boolean>;
  extraTracking: Record<Extra, boolean>;
  initialProfileId?: string;
  onSaved: (profileId: string) => Promise<void> | void;
  onClose: () => void;
};

const STANDARD = "standard";
const FLAGS: Flag[] = ["owned", "wearing", "foil", "foulborn", "vestigial"];
const ALIASES = new Map([["advanced fortress", "advancing fortress"]]);

function isFlag(value: string): value is Flag {
  return FLAGS.includes(value as Flag);
}
function style(color: string): CSSProperties | undefined {
  return color === "rainbow" ? undefined : ({ "--status-color": color } as CSSProperties);
}
function norm(value: string) {
  return value
    .normalize("NFKC")
    .replace(/[’‘]/g, "'")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function distance(a: string, b: string) {
  a = norm(a);
  b = norm(b);
  if (!a) return b.length;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i += 1) {
    const cur = [i];
    for (let j = 1; j <= b.length; j += 1) {
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}
function searchCatalogue(raw: string, preferredType: string, uniques: Unique[]) {
  const original = norm(raw);
  const query = ALIASES.get(original) ?? original;
  if (!query) return [];
  const tokens = query.split(" ").filter(Boolean);

  return uniques
    .map((unique) => {
      const name = norm(unique.name);
      const haystack = norm([unique.name, unique.baseType ?? "", unique.variantLabel ?? "", unique.itemType].join(" "));
      const d = distance(query, name);
      const exact = query === name;
      const substring = name.includes(query) || query.includes(name);
      const tokenHits = tokens.filter((token) => haystack.includes(token)).length;
      if (!exact && !substring && d > Math.max(3, Math.ceil(query.length * 0.35)) && tokenHits < Math.ceil(tokens.length / 2)) {
        return null;
      }
      let score = exact ? 0 : substring ? 10 + d : 30 + d * 2;
      score -= tokenHits * 4;
      if (unique.itemType === preferredType) score -= 6;
      return { unique, score };
    })
    .filter((entry): entry is { unique: Unique; score: number } => entry !== null)
    .sort((a, b) => a.score - b.score || a.unique.name.localeCompare(b.unique.name))
    .slice(0, 12)
    .map((entry) => entry.unique);
}

async function ensureSchema(db: Database) {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS import_review_progress (
      profile_id TEXT NOT NULL,
      import_id TEXT NOT NULL,
      decision TEXT NOT NULL,
      confirmed_at INTEGER NOT NULL,
      PRIMARY KEY (profile_id, import_id)
    )
  `);
}
async function loadItems(db: Database, profileId: string): Promise<ReviewItem[]> {
  await ensureSchema(db);
  const rows = await db.select<{
    import_id: string;
    name: string;
    item_type: string;
    import_status: string;
    reconciliation_status: "ambiguous" | "unmatched";
    candidates_json: string;
  }[]>(
    `
      SELECT reconciliation.import_id, imported.name, imported.item_type,
        imported.status AS import_status, reconciliation.status AS reconciliation_status,
        reconciliation.candidates_json
      FROM import_reconciliation reconciliation
      JOIN imported_collection imported ON imported.id = reconciliation.import_id
      LEFT JOIN import_review_progress progress
        ON progress.profile_id = ? AND progress.import_id = reconciliation.import_id
      WHERE reconciliation.status IN ('ambiguous', 'unmatched')
        AND progress.import_id IS NULL
      ORDER BY imported.name COLLATE NOCASE ASC, imported.item_type COLLATE NOCASE ASC
    `,
    [profileId],
  );

  return rows.map((row) => {
    let candidates: Candidate[] = [];
    try {
      candidates = JSON.parse(row.candidates_json) as Candidate[];
    } catch {
      candidates = [];
    }
    return {
      importId: row.import_id,
      name: row.name,
      itemType: row.item_type,
      importStatus: row.import_status,
      status: row.reconciliation_status,
      candidates,
    };
  });
}
async function loadProfile(db: Database, profileId: string) {
  const reviews = await db.select<{ unique_id: string; reviewed: number }[]>(
    `SELECT unique_id, reviewed FROM profile_collection_review WHERE profile_id = ?`,
    [profileId],
  );
  const tracking = await db.select<{ unique_id: string; flag: string }[]>(
    `SELECT unique_id, flag FROM profile_unique_tracking WHERE profile_id = ?`,
    [profileId],
  );
  const result: Record<string, ProfileStatus> = {};
  for (const row of reviews) result[row.unique_id] = { reviewed: row.reviewed === 1, flags: [] };
  for (const row of tracking) {
    if (!isFlag(row.flag)) continue;
    const current = result[row.unique_id] ?? { reviewed: false, flags: [] };
    if (!current.flags.includes(row.flag)) current.flags.push(row.flag);
    result[row.unique_id] = current;
  }
  return result;
}
async function saveReview(
  db: Database,
  profileId: string,
  draft: Draft,
) {
  await ensureSchema(db);

  const rows =
    Object.entries(draft).filter(
      ([, row]) =>
        row.decision !== "pending",
    );

  if (!rows.length) {
    throw new Error(
      "Review or ignore at least one row before confirming.",
    );
  }

  const merged =
    new Map<string, Flag[]>();

  for (const [, row] of rows) {
    if (row.decision !== "reviewed") {
      continue;
    }

    for (
      const [uniqueId, flags] of
      Object.entries(row.changes)
    ) {
      const next =
        Array.from(
          new Set(flags),
        ).sort();

      const old =
        merged.get(uniqueId);

      if (
        old &&
        JSON.stringify(old) !==
          JSON.stringify(next)
      ) {
        throw new Error(
          "Two import rows give the same catalogue variant different statuses.",
        );
      }

      merged.set(
        uniqueId,
        next,
      );
    }
  }

  const statements:
    SqliteTransactionStatement[] = [];

  for (
    const [uniqueId, flags] of
    merged
  ) {
    statements.push(
      {
        sql: `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        params: [
          profileId,
          uniqueId,
        ],
      },
      {
        sql: `
          DELETE FROM profile_unique_tracking
          WHERE
            profile_id = ?
            AND unique_id = ?
        `,
        params: [
          profileId,
          uniqueId,
        ],
      },
    );

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
          profileId,
          uniqueId,
          flag,
        ],
      });
    }
  }

  const confirmedAt =
    String(Date.now());

  for (
    const [importId, row] of
    rows
  ) {
    statements.push({
      sql: `
        INSERT OR REPLACE INTO import_review_progress (
          profile_id,
          import_id,
          decision,
          confirmed_at
        )
        VALUES (?, ?, ?, ?)
      `,
      params: [
        profileId,
        importId,
        row.decision,
        confirmedAt,
      ],
    });
  }

  await invoke(
    "execute_sqlite_transaction",
    {
      statements,
      commit: true,
    },
  );
}

function StatusButtons({
  reviewed,
  flags,
  unique,
  colors,
  rules,
  extraTracking,
  onChange,
}: {
  reviewed: boolean;
  flags: Flag[];
  unique: Unique;
  colors: Props["statusColors"];
  rules: Props["collectionRules"];
  extraTracking: Props["extraTracking"];
  onChange: (flags: Flag[]) => void;
}) {
  const missing = reviewed && !flags.some((flag) => rules[flag]);
  const toggle = (flag: Flag) =>
    onChange(flags.includes(flag) ? flags.filter((value) => value !== flag) : [...flags, flag]);
  const show = (edition: Extra) =>
    extraTracking[edition] && (flags.includes(edition) || unique.editionAvailability[edition] !== "unavailable");

  return (
    <div className="tracking-badges">
      <button type="button" className={`tracking-badge missing ${missing ? "active" : ""}`} style={style(colors.missing)} onClick={() => onChange([])}>
        Missing
      </button>
      <button type="button" className={`tracking-badge owned ${flags.includes("owned") ? "active" : ""}`} style={style(colors.owned)} onClick={() => toggle("owned")}>
        Owned
      </button>
      <button type="button" className={`tracking-badge wearing ${flags.includes("wearing") ? "active" : ""}`} style={style(colors.wearing)} onClick={() => toggle("wearing")}>
        Wearing
      </button>
      {(show("foil") || show("foulborn") || show("vestigial")) && <span className="tracking-divider" />}
      {show("foil") && (
        <button type="button" className={`tracking-badge foil ${flags.includes("foil") ? "active" : ""} ${colors.foil === "rainbow" ? "rainbow-status" : ""}`} style={style(colors.foil)} onClick={() => toggle("foil")}>
          Foil
        </button>
      )}
      {show("foulborn") && (
        <button type="button" className={`tracking-badge foulborn ${flags.includes("foulborn") ? "active" : ""}`} style={style(colors.foulborn)} onClick={() => toggle("foulborn")}>
          Foulborn
        </button>
      )}
      {show("vestigial") && (
        <button type="button" className={`tracking-badge vestigial ${flags.includes("vestigial") ? "active" : ""}`} style={style(colors.vestigial)} onClick={() => toggle("vestigial")}>
          Vestigial
        </button>
      )}
    </div>
  );
}

export default function ImportReviewModal({
  database,
  profiles,
  uniques,
  statusColors,
  collectionRules,
  extraTracking,
  initialProfileId,
  onSaved,
  onClose,
}: Props) {
  const [items, setItems] = useState<ReviewItem[]>([]);
  const [destinationId, setDestinationId] = useState(
    initialProfileId ?? STANDARD,
  );
  const [profileState, setProfileState] = useState<Record<string, ProfileStatus>>({});
  const [draft, setDraft] = useState<Draft>({});
  const [search, setSearch] = useState<Record<string, string>>({});
  const [busy, setBusy] = useState(true);
  const [message, setMessage] = useState("");
  const [pendingDestination, setPendingDestination] = useState<string | null>(null);
  const [expandReviewed, setExpandReviewed] = useState(false);
  const [expandedReviewed, setExpandedReviewed] =
    useState<Record<string, boolean>>({});

  const destination = profiles.find((profile) => profile.id === destinationId) ?? null;
  const decisionCount = Object.values(draft).filter((row) => row.decision !== "pending").length;
  const hasDraft = Object.values(draft).some(
    (row) => row.decision !== "pending" || Object.keys(row.changes).length > 0,
  );

  async function loadDestination(profileId: string, reloadItems: boolean) {
    const state = await loadProfile(database, profileId);
    if (reloadItems) {
      setItems(await loadItems(database, profileId));
      setSearch({});
    }
    setProfileState(state);
    setDestinationId(profileId);
  }

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        setBusy(true);
        setMessage("");
        await database.execute(`DELETE FROM import_manual_resolutions`).catch(() => undefined);
        await reconcileImportedCollection(database);
        const first = profiles.some(
          (profile) => profile.id === initialProfileId,
        )
          ? initialProfileId
          : profiles.some((profile) => profile.id === STANDARD)
            ? STANDARD
            : profiles[0]?.id;
        if (!first) throw new Error("No collection profile is available.");
        const [nextItems, state] = await Promise.all([loadItems(database, first), loadProfile(database, first)]);
        if (!cancelled) {
          setItems(nextItems);
          setProfileState(state);
          setDestinationId(first);
        }
      } catch (error) {
        if (!cancelled) setMessage(error instanceof Error ? error.message : String(error));
      } finally {
        if (!cancelled) setBusy(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [database, profiles, initialProfileId]);

  function row(importId: string): DraftRow {
    return draft[importId] ?? { decision: "pending", changes: {} };
  }
  function status(importId: string, uniqueId: string): ProfileStatus {
    const current = draft[importId];
    if (current && Object.prototype.hasOwnProperty.call(current.changes, uniqueId)) {
      return { reviewed: true, flags: current.changes[uniqueId] };
    }
    return profileState[uniqueId] ?? { reviewed: false, flags: [] };
  }
  function setDecision(importId: string, decision: Decision) {
    setDraft((current) => ({
      ...current,
      [importId]: {
        ...(current[importId] ?? { changes: {} }),
        decision,
      },
    }));

    setExpandedReviewed((current) => {
      const next = { ...current };
      delete next[importId];
      return next;
    });
  }
  function setFlags(item: ReviewItem, uniqueId: string, flags: Flag[]) {
    setDraft((current) => {
      const old = current[item.importId] ?? { decision: "pending" as const, changes: {} };
      const changes = { ...old.changes, [uniqueId]: flags };
      let decision: Decision = old.decision === "ignored" ? "pending" : old.decision;
      if (item.status === "unmatched") decision = "reviewed";
      else if (
        item.candidates.every(
          (candidate) =>
            Object.prototype.hasOwnProperty.call(changes, candidate.id) ||
            (profileState[candidate.id]?.reviewed ?? false),
        )
      ) {
        decision = "reviewed";
      }
      return { ...current, [item.importId]: { decision, changes } };
    });
  }
  function changeDestination(profileId: string) {
    if (profileId === destinationId) return;
    if (hasDraft) {
      setPendingDestination(profileId);
      return;
    }
    setBusy(true);
    void loadDestination(profileId, true)
      .catch((error) => setMessage(error instanceof Error ? error.message : String(error)))
      .finally(() => setBusy(false));
  }
  async function confirmDestinationChange() {
    if (!pendingDestination) return;
    try {
      setBusy(true);
      await loadDestination(pendingDestination, false);
      setPendingDestination(null);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }
  async function confirm() {
    try {
      setBusy(true);
      setMessage("");
      await saveReview(database, destinationId, draft);
      await onSaved(destinationId);
      onClose();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function candidate(item: ReviewItem, unique: Unique) {
    const current = status(item.importId, unique.id);
    const changed = Object.prototype.hasOwnProperty.call(row(item.importId).changes, unique.id);
    return (
      <div key={unique.id} style={{ border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: 12 }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12 }}>
          <div>
            <strong>{unique.name}</strong>
            <div className="unique-variant-meta">
              {[unique.baseType, unique.variantLabel, unique.itemType].filter(Boolean).join(" • ")}
            </div>
          </div>
          <span className={current.reviewed ? "legacy-badge" : "unreviewed-badge"}>
            {changed ? "DRAFT CHANGE" : current.reviewed ? "CURRENTLY REVIEWED" : "UNREVIEWED"}
          </span>
        </div>
        <div style={{ marginTop: 10 }}>
          <StatusButtons
            reviewed={current.reviewed}
            flags={current.flags}
            unique={unique}
            colors={statusColors}
            rules={collectionRules}
            extraTracking={extraTracking}
            onChange={(flags) => setFlags(item, unique.id, flags)}
          />
        </div>
      </div>
    );
  }

  return (
    <>
      <div className="catalogue-update-overlay">
        <section className="catalogue-update-modal" style={{ width: "min(980px, calc(100vw - 40px))", maxWidth: 980, maxHeight: "90vh", display: "flex", flexDirection: "column" }}>
          <div className="catalogue-update-heading">
            <span className="catalogue-update-kicker">IMPORT REVIEW</span>
            <h2>Resolve Imported Collection</h2>
            <p>Review everything in one batch. Nothing is saved until you confirm.</p>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
            <div className="filter-field">
              <label htmlFor="import-review-destination">Save review to</label>
              <select id="import-review-destination" value={destinationId} disabled={busy} onChange={(event) => changeDestination(event.target.value)}>
                {profiles.map((profile) => (
                  <option value={profile.id} key={profile.id}>{profile.name}</option>
                ))}
              </select>
            </div>
            <div className="hotkey-test-status ready">
              <strong>Reviewing for</strong>
              <span>{destination?.name ?? "Standard"} • {decisionCount} of {items.length} reviewed/ignored</span>
            </div>
          </div>

          {decisionCount > 0 && (
            <div
              className="catalogue-update-actions"
              style={{
                justifyContent: "flex-start",
                marginBottom: 14,
              }}
            >
              <button
                type="button"
                className="catalogue-update-secondary"
                disabled={busy}
                onClick={() => {
                  if (expandReviewed) {
                    setExpandedReviewed({});
                  }

                  setExpandReviewed(
                    (current) => !current,
                  );
                }}
              >
                {expandReviewed
                  ? "Collapse Reviewed"
                  : "Expand Reviewed"}
              </button>
            </div>
          )}

          {message && <p className="catalogue-check-status">{message}</p>}

          <div style={{ overflowY: "auto", display: "grid", gap: 12, minHeight: 0, paddingRight: 4 }}>
            {busy && !items.length ? (
              <p className="settings-help">Loading import review...</p>
            ) : !items.length ? (
              <p className="settings-help">Nothing left to review for {destination?.name ?? "this collection"}.</p>
            ) : (
              items.map((item) => {
                const currentRow = row(item.importId);
                const candidates =
                  item.status === "ambiguous"
                    ? item.candidates.map((value) => uniques.find((unique) => unique.id === value.id)).filter((unique): unique is Unique => unique !== undefined)
                    : [];
                const searchTerm = search[item.importId] ?? item.name;
                const results = item.status === "unmatched" ? searchCatalogue(searchTerm, item.itemType, uniques) : [];
                const isCollapsed =
                  currentRow.decision !== "pending" &&
                  !expandReviewed &&
                  !expandedReviewed[item.importId];

                return (
                  <div key={item.importId} style={{ border: currentRow.decision === "reviewed" ? "1px solid rgba(121,186,104,0.55)" : "1px solid rgba(255,255,255,0.1)", borderRadius: 12, padding: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 12 }}>
                      <div>
                        <h3 style={{ margin: 0 }}>{item.name}</h3>
                        <p className="settings-help" style={{ marginBottom: 0 }}>Spreadsheet: {item.importStatus} • {item.itemType}</p>
                      </div>
                      <span className={currentRow.decision === "pending" ? "unreviewed-badge" : "legacy-badge"}>
                        {currentRow.decision === "reviewed" ? "REVIEWED" : currentRow.decision === "ignored" ? "IGNORED" : "NEEDS REVIEW"}
                      </span>
                    </div>

                    {isCollapsed ? (
                      <div
                        className="catalogue-update-actions"
                        style={{
                          justifyContent: "space-between",
                          marginTop: 12,
                        }}
                      >
                        <span
                          className="settings-help"
                          style={{ margin: 0 }}
                        >
                          {currentRow.decision === "ignored"
                            ? "This spreadsheet row will be ignored."
                            : "Reviewed — expand to inspect or change it before confirming."}
                        </span>

                        <button
                          type="button"
                          className="catalogue-update-secondary"
                          disabled={busy}
                          onClick={() =>
                            setExpandedReviewed((current) => ({
                              ...current,
                              [item.importId]: true,
                            }))
                          }
                        >
                          Expand
                        </button>
                      </div>
                    ) : (
                      <>
                        <div
                          style={{
                            display: "grid",
                            gap: 8,
                            marginTop: 12,
                          }}
                        >
                          {item.status === "ambiguous" ? (
                            candidates.length ? (
                              candidates.map((unique) =>
                                candidate(item, unique),
                              )
                            ) : (
                              <p className="settings-help">
                                No saved candidates remain in the catalogue.
                              </p>
                            )
                          ) : (
                            <>
                              <p
                                className="settings-help"
                                style={{ marginBottom: 0 }}
                              >
                                No safe automatic match. Search uses fuzzy
                                matching and known aliases.
                              </p>

                              <div className="search-input-wrap">
                                <input
                                  type="text"
                                  value={searchTerm}
                                  onChange={(event) =>
                                    setSearch((old) => ({
                                      ...old,
                                      [item.importId]:
                                        event.target.value,
                                    }))
                                  }
                                />

                                {searchTerm && (
                                  <button
                                    type="button"
                                    className="search-clear-button"
                                    aria-label="Clear search"
                                    onClick={() =>
                                      setSearch((old) => ({
                                        ...old,
                                        [item.importId]: "",
                                      }))
                                    }
                                  >
                                    ×
                                  </button>
                                )}
                              </div>

                              {results.length ? (
                                results.map((unique) =>
                                  candidate(item, unique),
                                )
                              ) : (
                                <p className="settings-help">
                                  No useful suggestions. Try another spelling
                                  or leave it for later.
                                </p>
                              )}
                            </>
                          )}
                        </div>

                        <div
                          className="catalogue-update-actions"
                          style={{
                            justifyContent: "flex-start",
                            marginTop: 12,
                          }}
                        >
                          <button
                            type="button"
                            className="catalogue-update-secondary"
                            disabled={busy}
                            onClick={() =>
                              setDecision(
                                item.importId,
                                "reviewed",
                              )
                            }
                          >
                            Mark Reviewed
                          </button>

                          <button
                            type="button"
                            className="catalogue-update-secondary"
                            disabled={busy}
                            onClick={() =>
                              setDecision(
                                item.importId,
                                "ignored",
                              )
                            }
                          >
                            Ignore This Import
                          </button>

                          {currentRow.decision !== "pending" && (
                            <button
                              type="button"
                              className="catalogue-update-secondary"
                              disabled={busy}
                              onClick={() =>
                                setDecision(
                                  item.importId,
                                  "pending",
                                )
                              }
                            >
                              Mark Unreviewed
                            </button>
                          )}

                          {currentRow.decision !== "pending" &&
                            !expandReviewed && (
                              <button
                                type="button"
                                className="catalogue-update-secondary"
                                disabled={busy}
                                onClick={() =>
                                  setExpandedReviewed(
                                    (current) => ({
                                      ...current,
                                      [item.importId]: false,
                                    }),
                                  )
                                }
                              >
                                Collapse
                              </button>
                            )}
                        </div>
                      </>
                    )}
                  </div>
                );
              })
            )}
          </div>

          <div className="catalogue-update-actions" style={{ marginTop: 16 }}>
            <button type="button" className="catalogue-update-secondary" disabled={busy} onClick={onClose}>Cancel Review</button>
            <button type="button" className="catalogue-update-primary" disabled={busy || decisionCount === 0} onClick={() => void confirm()}>
              {busy ? "Saving Review..." : `Confirm ${decisionCount} ${decisionCount === 1 ? "Change" : "Changes"} to ${destination?.name ?? "Standard"}`}
            </button>
          </div>
        </section>
      </div>

      {pendingDestination && (
        <div className="catalogue-update-overlay">
          <section className="catalogue-update-modal">
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">CHANGE REVIEW DESTINATION</span>
              <h2>Switch to {profiles.find((profile) => profile.id === pendingDestination)?.name ?? "that collection"}?</h2>
              <p>Your draft is not saved. If you switch, all current review choices will be applied to the newly selected collection when you confirm.</p>
            </div>
            <div className="catalogue-update-actions">
              <button type="button" className="catalogue-update-secondary" disabled={busy} onClick={() => setPendingDestination(null)}>Keep {destination?.name ?? "Current"}</button>
              <button type="button" className="catalogue-update-primary" disabled={busy} onClick={() => void confirmDestinationChange()}>Switch Collection</button>
            </div>
          </section>
        </div>
      )}
    </>
  );
}
