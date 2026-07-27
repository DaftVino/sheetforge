# ADR 0002: Decouple the contract version from the stored-format schema marker

<!-- File: docs/adr/0002-contract-version-vs-schema-marker.md — one page max. Never edited after acceptance; reversals are a new ADR. -->

**Status:** Rejected
**Date:** 2026-07-27

**Rejected by the owner on 2026-07-27**, on a fact this draft did not have: **there is no installed base.** The bundle exists only in test workbooks, so the false refusal this ADR was written to prevent costs nothing today. One version number is simpler than two, and simpler wins while it is free. `SCHEMA_VER` therefore moves to `"v2.0"` with the contract, and contract §10 records both the decision and the asymmetry it creates.

**Revisit before the next breaking bump.** The analysis below does not stop being true — it stops being *expensive*. Once real deployments exist, refusing every stored config to express a change that touched no stored byte is a genuine cost, and the decision should be retaken on the merits set out here rather than re-derived from scratch. That is why this stays on the record as Rejected rather than being deleted.

## Context

`spec/contract.md` §10 versions the contract with SemVer and classifies "renaming or
removing a `key` or action `id`; changing … an action's semantics" as **breaking**,
bumping the **major**.

Issue #5 (Part B) withdraws `clean_keycols` and `clean_hasheader` and changes
`clean.dupes` to take those two answers as run-time inputs. By the letter of §10 this
is breaking, and the contract goes to **v2.0.0**.

Separately, `Core.bas` carries `Private Const SCHEMA_VER As String = "v1.2"` — a
machine-owned marker written into the hidden id column of the Action catalog's key row
in `_Config`. `MarkerVerdict` compares its major against `CurrentSchemaMajor()`:

- equal major → `same-major`, read normally
- different major → `other-major`, which `GateConfigMigration` maps to `unknown` and
  which makes `Create Setup Tab` **refuse outright**, telling the user to re-import
  that version's bundle

That refusal is correct and deliberate: an old build must never rewrite a newer format
it cannot read. It was verified by hand in the v0.4.0 gate (Task 10, Step 7b).

**The collision.** `SCHEMA_VER`'s spelling tracks the contract version by convention —
contract v1.2.0, marker `v1.2`. If that convention is honoured through a contract
v2.0.0, the marker becomes `v2.x`, every existing v0.4.0 `_Config` reads as
`other-major`, and **every current deployment is refused on its next `Create Setup
Tab`** — for a change that alters no stored layout whatsoever. The Action catalog block
is byte-identical before and after Part B: same three columns, same key row, same
per-row ids. The only thing that changed is that two field rows are no longer declared,
which the reader already ignores.

## Decision proposed — and NOT taken

*What follows is the rejected proposal, kept intact as the record of what was weighed.
The decision actually taken is at the top of this file.*

State in §10 that the two version numbers are **independent**, and say what each one
describes:

- **The contract version** describes the *published surface*: keys, action ids, kinds,
  encodings, action semantics. Part B takes it to **v2.0.0**.
- **`SCHEMA_VER`** describes the *stored block layout* — the physical shape of the
  Action catalog block that a migration has to read. It bumps **only** when that shape
  changes, and it stays at **`v1.2`** through Part B.

A withdrawn field does not change the stored layout, so it does not move the marker.
Adding, removing, or reordering a *column* of the catalog block does.

## Alternatives considered

- **Bump both to major 2.** Rejected: it refuses every existing deployment to express a
  change none of them are affected by, and the recovery it names ("re-import that
  version's bundle") would be actively wrong, since this build reads those files fine.
- **Call Part B additive and stay on v1.x.** Rejected: it is a removal of two keys and
  a semantic change to an action. Calling that additive would make §10's own
  classification meaningless, and the next reader could not trust it.
- **Keep them coupled and add a migration that rewrites old markers.** Rejected: it
  adds a destructive rewrite path over `_Config` to work around a version number, and
  the fail-closed guarantee this repo just spent a gate proving is worth more than the
  tidiness of one convention.

## Consequences

- §10 gains an explicit statement that the two numbers are independent, with the rule
  for when each moves. Until now the coupling was convention, undocumented — which is
  exactly why it was able to collide.
- Every port must implement the same split, or a port could refuse configs the
  reference build accepts.
- The `v1.2` marker outlives contract v1.x, so the spelling no longer hints at the
  contract version. That is a readability cost and is the strongest argument against;
  it is accepted because the alternative is a false refusal.
- **Open sub-question for the owner:** should `SCHEMA_VER` be respelled to something
  that cannot be mistaken for a contract version (e.g. `layout-3`) as part of this? That
  is a stored-format change and would itself trip the migration, so it can only ride
  along with a real layout bump — not with this one.
