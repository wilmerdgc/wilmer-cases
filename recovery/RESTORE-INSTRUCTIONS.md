# Recover your erased cloud data

## What happened

Your case data lives in a Firebase/Firestore cloud database that all your
devices sync to. An **old device running an old version of the app** signed in
and saved. That old version pre-dates the app's per-record merge protection, so
instead of merging it did a whole-document overwrite — replacing the cloud copy
with its stale snapshot and wiping the work you'd done since.

## What we have

A backup bot snapshots your cloud data into this repo every day. The last clean
snapshot before the overwrite is:

- **`backups/2026-06-27.json`** — **741 records**, all settings, and **147 case
  files** (saved 2026-06-26). This is your full data.

A copy is included here as `recovery/last-good-snapshot-2026-06-27.json`.

## How to restore (recommended — one click)

Open **`recovery/restore.html`** in your normal web browser (double-click the
file, or host it). It:

1. Signs in to your cloud the same way the app does.
2. Shows how much data it will restore (741 records / 147 case files).
3. On **"Restore my data to the cloud"**, pushes the snapshot back using the
   **same safe per-record merge the app itself uses.**

**It is safe to run.** The merge only *adds back* records that were erased and
keeps the **newest** version of every record. It can never overwrite anything in
the cloud that is newer than the snapshot, and it never deletes. If anything
fails midway, just click Restore again — it is repeatable.

After it finishes: open the app on your normal device and tap **Pull from
cloud** (or just reload). Every other device will sync automatically.

> Tip: After restoring, **don't open the app on the old device again** until it
> has loaded the current data — an old app version can repeat the overwrite.
> Reload it / clear its old data first.

## Fallback (in-app Import)

If you'd rather not use the restore page, in the app open **Import** and choose
`recovery/last-good-snapshot-2026-06-27.json`, then pick **Replace**. This
restores your records, case files, paralegals, event types and case types, then
re-syncs to the cloud. (The one-click restore page above also brings back
judges, attorney emails, signatures, venues and firm info, so prefer it if you
can.)

## Why this won't happen again

The current app version already merges per-record (`_updatedAt` wins) instead of
overwriting, so a normal up-to-date device can no longer clobber the cloud. The
remaining risk is an *old* device/tab still running an old build — close or
refresh those so they pick up the current version.
