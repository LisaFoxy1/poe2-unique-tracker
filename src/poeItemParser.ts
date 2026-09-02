import Database from "@tauri-apps/plugin-sql";

export type ItemEdition = "normal" | "foulborn" | "vestigial";

export type ParsedPoeItem = {
  itemClass: string | null;
  rarity: string | null;
  name: string | null;
  baseType: string | null;
  canonicalName: string | null;
  canonicalBaseType: string | null;
  edition: ItemEdition;
  rawText: string;
};

export type ItemIdentificationResult =
  | {
      status: "matched";
      parsed: ParsedPoeItem;
      uniqueId: string;
      familyId: string;
      name: string;
      baseType: string | null;
      itemType: string;
      variantLabel: string | null;
      edition: ItemEdition;
      matchedBy:
        | "special-rule"
        | "variant-label"
        | "name-and-base";
    }
  | {
      status: "ambiguous";
      parsed: ParsedPoeItem;
      message: string;
      candidates: Array<{
        uniqueId: string;
        name: string;
        baseType: string | null;
        variantLabel: string | null;
      }>;
    }
  | {
      status: "not-found" | "not-unique" | "invalid";
      parsed: ParsedPoeItem;
      message: string;
    };

type CandidateRow = {
  id: string;
  family_id: string;
  name: string;
  base_type: string | null;
  item_type: string;
  variant_label: string | null;
};

type RuleRow = CandidateRow & {
  match_type: string;
  match_text: string;
  priority: number;
};

function cleanLines(rawText: string) {
  return rawText
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.trim());
}

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

function cleanIdentity(value: string | null) {
  if (!value) {
    return null;
  }

  return value
    .normalize("NFKC")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/\s+/g, " ")
    .trim();
}

function detectEdition(
  rawText: string,
  name: string | null,
  baseType: string | null,
) {
  const normalizedRaw = normalize(rawText);

  const isFoulborn =
    normalize(name).startsWith("foulborn ") ||
    normalizedRaw.includes("{ foulborn unique modifier }");

  const isVestigial =
    normalize(baseType).startsWith("vestigial ") ||
    normalizedRaw.includes("{ vestigial implicit modifier");

  let edition: ItemEdition = "normal";

  if (isFoulborn) {
    edition = "foulborn";
  } else if (isVestigial) {
    edition = "vestigial";
  }

  let canonicalName = name;
  let canonicalBaseType = baseType;

  if (edition === "foulborn" && canonicalName) {
    canonicalName = canonicalName.replace(/^Foulborn\s+/i, "").trim();
  }

  if (edition === "vestigial" && canonicalBaseType) {
    canonicalBaseType = canonicalBaseType
      .replace(/^Vestigial\s+/i, "")
      .trim();
  }

  return {
    edition,
    canonicalName: cleanIdentity(canonicalName),
    canonicalBaseType: cleanIdentity(canonicalBaseType),
  };
}

export function parsePoeItemText(rawText: string): ParsedPoeItem {
  const lines = cleanLines(rawText);

  const itemClassLine = lines.find((line) =>
    line.startsWith("Item Class:"),
  );

  const rarityIndex = lines.findIndex((line) =>
    line.startsWith("Rarity:"),
  );

  const rarity =
    rarityIndex >= 0
      ? lines[rarityIndex].slice("Rarity:".length).trim() || null
      : null;

  let name: string | null = null;
  let baseType: string | null = null;

  if (rarityIndex >= 0) {
    const afterRarity = lines
      .slice(rarityIndex + 1)
      .filter((line) => line !== "");

    if (afterRarity.length > 0 && afterRarity[0] !== "--------") {
      name = afterRarity[0];
    }

    if (
      afterRarity.length > 1 &&
      afterRarity[1] !== "--------" &&
      !afterRarity[1].startsWith("Item Level:")
    ) {
      baseType = afterRarity[1];
    }
  }

  const cleanName = cleanIdentity(name);
  const cleanBaseType = cleanIdentity(baseType);

  const {
    edition,
    canonicalName,
    canonicalBaseType,
  } = detectEdition(
    rawText,
    cleanName,
    cleanBaseType,
  );

  return {
    itemClass: itemClassLine
      ? cleanIdentity(
          itemClassLine.slice("Item Class:".length).trim() || null,
        )
      : null,
    rarity: cleanIdentity(rarity),
    name: cleanName,
    baseType: cleanBaseType,
    canonicalName,
    canonicalBaseType,
    edition,
    rawText,
  };
}

function toCandidate(row: CandidateRow) {
  return {
    uniqueId: row.id,
    name: row.name,
    baseType: row.base_type,
    variantLabel: row.variant_label,
  };
}

const VARIANT_LABEL_NOISE_WORDS =
  new Set([
    "current",
    "legacy",
    "variant",
    "elder",
    "uber",
    "curse",
    "curses",
    "normal",
    "version",
    "ring",
  ]);

function getVariantLabelTokens(
  label: string | null,
) {
  const normalizedLabel =
    normalize(label);

  if (!normalizedLabel) {
    return [] as string[];
  }

  const isNonSynthesised =
    normalizedLabel.includes(
      "non-synthesised",
    ) ||
    normalizedLabel.includes(
      "non-synthesized",
    );

  return Array.from(
    new Set(
      (
        normalizedLabel.match(
          /[a-z0-9]+/g,
        ) ?? []
      ).filter((token) => {
        if (
          VARIANT_LABEL_NOISE_WORDS.has(
            token,
          )
        ) {
          return false;
        }

        if (
          isNonSynthesised &&
          (
            token === "synthesised" ||
            token === "synthesized"
          )
        ) {
          return false;
        }

        /*
         * Numbers are handled by explicit count
         * checks below. Bare numbers are far too
         * common in PoE item text to be safe
         * evidence by themselves.
         */
        return token.length >= 4;
      }),
    ),
  );
}

function countOccurrences(
  text: string,
  search: string,
) {
  if (!search) {
    return 0;
  }

  let count = 0;
  let position = 0;

  while (true) {
    const found =
      text.indexOf(
        search,
        position,
      );

    if (found === -1) {
      break;
    }

    count += 1;

    position =
      found + search.length;
  }

  return count;
}

type VariantEvidence = {
  candidate: CandidateRow;
  score: number;
  strong: boolean;
  matchedTokens: number;
  tokens: string[];
};

function scoreVariantEvidence(
  candidate: CandidateRow,
  normalizedRaw: string,
): VariantEvidence {
  const label =
    normalize(
      candidate.variant_label,
    );

  const tokens =
    getVariantLabelTokens(
      candidate.variant_label,
    );

  let score = 0;
  let strong = false;

  if (!label) {
    return {
      candidate,
      score,
      strong,
      matchedTokens: 0,
      tokens,
    };
  }

  /*
   * A long label appearing literally in the
   * copied item is extremely strong evidence.
   *
   * Do not do this for tiny labels such as
   * "Fire" because those words can occur for
   * unrelated reasons.
   */
  if (
    label.length >= 10 &&
    normalizedRaw.includes(label)
  ) {
    score += 30;
    strong = true;
  }

  /*
   * Abyss uniques often differ only by their
   * number of Abyssal Sockets.
   */
  const labelAbyssSockets =
    label.match(
      /\b(\d+)\s+abyssal sockets?\b/,
    );

  const rawAbyssSockets =
    normalizedRaw.match(
      /\bhas\s+(\d+)\s+abyssal sockets?\b/,
    );

  if (
    labelAbyssSockets &&
    rawAbyssSockets
  ) {
    if (
      labelAbyssSockets[1] !==
      rawAbyssSockets[1]
    ) {
      return {
        candidate,
        score: -1,
        strong: false,
        matchedTokens: 0,
        tokens,
      };
    }

    score += 40;
    strong = true;
  }

  /*
   * Old Synthesis variants are directly marked
   * in copied item text.
   */
  if (
    label.includes("fractured") &&
    normalizedRaw.includes(
      "fractured",
    )
  ) {
    score += 40;
    strong = true;
  }

  const wantsSynthesised =
    (
      label.includes(
        "synthesised",
      ) ||
      label.includes(
        "synthesized",
      )
    ) &&
    !label.includes(
      "non-synthesised",
    ) &&
    !label.includes(
      "non-synthesized",
    );

  if (
    wantsSynthesised &&
    (
      normalizedRaw.includes(
        "synthesised",
      ) ||
      normalizedRaw.includes(
        "synthesized",
      )
    )
  ) {
    score += 40;
    strong = true;
  }

  /*
   * Thread of Hope has one genuinely distinct
   * Massive Ring variant.
   */
  if (
    label.includes(
      "massive ring",
    ) &&
    normalizedRaw.includes(
      "massive ring",
    )
  ) {
    score += 40;
    strong = true;
  }

  /*
   * The ordinary Thread of Hope can have one
   * of several non-Massive ring sizes.
   */
  if (
    label.includes(
      "variable ring",
    ) &&
    (
      normalizedRaw.includes(
        "small ring",
      ) ||
      normalizedRaw.includes(
        "medium ring",
      ) ||
      normalizedRaw.includes(
        "large ring",
      ) ||
      normalizedRaw.includes(
        "very large ring",
      )
    )
  ) {
    score += 40;
    strong = true;
  }

  /*
   * Old Beachhead variants are separated by
   * their Map Tier.
   */
  const labelTier =
    label.match(
      /\btier\s+(\d+)\b/,
    );

  const rawTier =
    normalizedRaw.match(
      /\bmap tier:\s*(\d+)\b/,
    );

  if (
    labelTier &&
    rawTier
  ) {
    if (
      labelTier[1] !==
      rawTier[1]
    ) {
      return {
        candidate,
        score: -1,
        strong: false,
        matchedTokens: 0,
        tokens,
      };
    }

    score += 40;
    strong = true;
  }

  /*
   * Impresence's stronger versions have two
   * reservation-free curses instead of one.
   */
  const labelCurseCount =
    label.match(
      /\b([12])\s+curses?\b/,
    );

  if (labelCurseCount) {
    const rawCurseCount =
      countOccurrences(
        normalizedRaw,
        "has no reservation if cast as an aura",
      );

    if (rawCurseCount > 0) {
      if (
        Number(
          labelCurseCount[1],
        ) !== rawCurseCount
      ) {
        return {
          candidate,
          score: -1,
          strong: false,
          matchedTokens: 0,
          tokens,
        };
      }

      score += 40;
      strong = true;
    }
  }

  /*
   * If a Cane is copied while its Veiled mods
   * are still visible, the count can distinguish
   * the three- and four-Veiled configurations.
   *
   * Fully unveiled Canes may still require a
   * dedicated rule later.
   */
  const labelVeiledCount =
    label.match(
      /\b([34])\s+veiled\b/,
    );

  if (labelVeiledCount) {
    const rawVeiledCount =
      countOccurrences(
        normalizedRaw,
        "veiled",
      );

    if (rawVeiledCount > 0) {
      if (
        Number(
          labelVeiledCount[1],
        ) !== rawVeiledCount
      ) {
        return {
          candidate,
          score: -1,
          strong: false,
          matchedTokens: 0,
          tokens,
        };
      }

      score += 30;
      strong = true;
    }
  }

  /*
   * Finally use descriptive Wiki/curated labels.
   *
   * Examples:
   * - Cold + Lightning Resistance
   * - Strength
   * - Fire
   * - Rhoa
   * - Armour + Energy Shield
   *
   * This is only a score. The caller still
   * requires one clearly better candidate.
   */
  const matchedTokens =
    tokens.filter((token) =>
      normalizedRaw.includes(token),
    ).length;

  score += matchedTokens;

  return {
    candidate,
    score,
    strong,
    matchedTokens,
    tokens,
  };
}

function tryMatchVariantByLabel(
  candidates: CandidateRow[],
  rawText: string,
) {
  const normalizedRaw =
    normalize(rawText);

  const evidence =
    candidates
      .map((candidate) =>
        scoreVariantEvidence(
          candidate,
          normalizedRaw,
        ),
      )
      .filter(
        (entry) =>
          entry.score > 0,
      )
      .sort(
        (a, b) =>
          b.score - a.score,
      );

  if (evidence.length === 0) {
    return null;
  }

  const best = evidence[0];
  const second =
    evidence[1] ?? null;

  /*
   * A tie means we learned nothing useful.
   */
  if (
    second &&
    best.score === second.score
  ) {
    return null;
  }

  /*
   * Explicit structural evidence such as
   * socket count / Fractured / Massive Ring
   * is enough on its own.
   */
  if (best.strong) {
    return best.candidate;
  }

  /*
   * Two or more descriptive label words matching
   * one candidate gives us a conservative winner.
   */
  if (
    best.matchedTokens >= 2
  ) {
    return best.candidate;
  }

  /*
   * Allow simple one-word variant families such as:
   *
   * Fire / Cold / Lightning
   * Strength / Dexterity / Intelligence
   * Rhoa / Snake / Ursa
   *
   * But only when that word belongs to exactly one
   * candidate and no competing candidate scored.
   */
  if (
    best.matchedTokens === 1 &&
    best.tokens.length === 1 &&
    !second
  ) {
    const token =
      best.tokens[0];

    const candidatesUsingToken =
      candidates.filter(
        (candidate) =>
          getVariantLabelTokens(
            candidate.variant_label,
          ).includes(token),
      );

    if (
      candidatesUsingToken.length === 1
    ) {
      return best.candidate;
    }
  }

  return null;
}

export async function identifyUniqueFromClipboard(
  db: Database,
  rawText: string,
): Promise<ItemIdentificationResult> {
  const parsed = parsePoeItemText(rawText);

  if (!parsed.canonicalName || !parsed.rarity) {
    return {
      status: "invalid",
      parsed,
      message:
        "Could not find a PoE item name and rarity in the pasted text.",
    };
  }

  if (normalize(parsed.rarity) !== "unique") {
    return {
      status: "not-unique",
      parsed,
      message: `The pasted item is ${parsed.rarity}, not Unique.`,
    };
  }

  const allSpecialRules = await db.select<RuleRow[]>(`
    SELECT
      v.id,
      v.family_id,
      v.name,
      v.base_type,
      v.item_type,
      v.variant_label,
      r.match_type,
      r.match_text,
      r.priority
    FROM unique_variant_rules r
    JOIN unique_variants v
      ON v.id = r.variant_id
    ORDER BY r.priority DESC
  `);

  const specialRules = allSpecialRules.filter(
    (rule) => normalize(rule.name) === normalize(parsed.canonicalName),
  );

  const normalizedRaw = normalize(rawText);

  const matchingRules = specialRules.filter((rule) => {
    if (rule.match_type === "contains") {
      return normalizedRaw.includes(normalize(rule.match_text));
    }

    if (rule.match_type === "contains-all") {
      const requiredParts = rule.match_text
        .split("|||")
        .map((part) => normalize(part))
        .filter(Boolean);

      return (
        requiredParts.length > 0 &&
        requiredParts.every((part) =>
          normalizedRaw.includes(part),
        )
      );
    }

    if (rule.match_type === "exact") {
      return normalizedRaw === normalize(rule.match_text);
    }

    return false;
  });

  if (matchingRules.length > 0) {
    const highestPriority = Math.max(
      ...matchingRules.map((rule) => rule.priority),
    );

    const highestPriorityRules = matchingRules.filter(
      (rule) => rule.priority === highestPriority,
    );

    if (highestPriorityRules.length === 1) {
      const match = highestPriorityRules[0];

      return {
        status: "matched",
        parsed,
        uniqueId: match.id,
        familyId: match.family_id,
        name: match.name,
        baseType: match.base_type,
        itemType: match.item_type,
        variantLabel: match.variant_label,
        edition: parsed.edition,
        matchedBy: "special-rule",
      };
    }

    return {
      status: "ambiguous",
      parsed,
      message:
        "More than one equally specific special variant rule matched this item. This catalogue entry needs a consistency review.",
      candidates: highestPriorityRules.map(toCandidate),
    };
  }

  const allCandidates = await db.select<CandidateRow[]>(`
    SELECT
      id,
      family_id,
      name,
      base_type,
      item_type,
      variant_label
    FROM unique_variants
  `);

  const sameName = allCandidates
    .filter(
      (candidate) =>
        normalize(candidate.name) === normalize(parsed.canonicalName),
    )
    .sort((a, b) => {
      const aExactBase =
        parsed.canonicalBaseType &&
        normalize(a.base_type) === normalize(parsed.canonicalBaseType)
          ? 0
          : 1;

      const bExactBase =
        parsed.canonicalBaseType &&
        normalize(b.base_type) === normalize(parsed.canonicalBaseType)
          ? 0
          : 1;

      if (aExactBase !== bExactBase) {
        return aExactBase - bExactBase;
      }

      return (a.variant_label ?? "").localeCompare(
        b.variant_label ?? "",
      );
    });

  if (sameName.length === 0) {
    return {
      status: "not-found",
      parsed,
      message: `${parsed.canonicalName} was not found in the local catalogue.`,
    };
  }

  const exactBase = parsed.canonicalBaseType
    ? sameName.filter(
        (candidate) =>
          normalize(candidate.base_type) ===
          normalize(parsed.canonicalBaseType),
      )
    : [];

  const usableCandidates =
    exactBase.length > 0 ? exactBase : sameName;

  if (usableCandidates.length === 1) {
    const match = usableCandidates[0];

    return {
      status: "matched",
      parsed,
      uniqueId: match.id,
      familyId: match.family_id,
      name: match.name,
      baseType: match.base_type,
      itemType: match.item_type,
      variantLabel: match.variant_label,
      edition: parsed.edition,
      matchedBy: "name-and-base",
    };
  }

  const variantLabelMatch =
  tryMatchVariantByLabel(
    usableCandidates,
    rawText,
  );

if (variantLabelMatch) {
  return {
    status: "matched",
    parsed,
    uniqueId:
      variantLabelMatch.id,
    familyId:
      variantLabelMatch.family_id,
    name:
      variantLabelMatch.name,
    baseType:
      variantLabelMatch.base_type,
    itemType:
      variantLabelMatch.item_type,
    variantLabel:
      variantLabelMatch.variant_label,
    edition: parsed.edition,
    matchedBy: "variant-label",
  };
}

  return {
    status: "ambiguous",
    parsed,
    message:
      "The name/base matches multiple collectible variants, but the pasted item did not contain enough information to choose one safely.",
    candidates: usableCandidates.map(toCandidate),
  };
}