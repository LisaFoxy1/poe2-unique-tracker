import { useEffect, useState } from "react";
import { getVersion } from "@tauri-apps/api/app";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import {
  check,
  type Update,
} from "@tauri-apps/plugin-updater";

const PENDING_UPDATE_KEY =
  "poe2-collector-pending-update";

type UpdateNotice = {
  version: string;
  notes: string;
};

async function revealMainWindow() {
  const mainWindow =
    await WebviewWindow.getByLabel("main");

  if (!mainWindow) {
    return;
  }

  await mainWindow.show();
  await mainWindow.unminimize();
  await mainWindow.setFocus();
}

export default function AppUpdater() {
  const [availableUpdate, setAvailableUpdate] =
    useState<Update | null>(null);

  const [updateOpen, setUpdateOpen] =
    useState(false);

  const [downloading, setDownloading] =
    useState(false);

  const [downloaded, setDownloaded] =
    useState(false);

  const [installing, setInstalling] =
    useState(false);

  const [downloadProgress, setDownloadProgress] =
    useState<number | null>(null);

  const [updateError, setUpdateError] =
    useState("");

  const [installedNotice, setInstalledNotice] =
    useState<UpdateNotice | null>(null);

  useEffect(() => {
    if (import.meta.env.DEV) {
      return;
    }

    let cancelled = false;

    const timer = window.setTimeout(() => {
      void checkForAppUpdate();
    }, 1500);

    async function checkForAppUpdate() {
      try {
        const currentVersion = await getVersion();
        const pendingRaw =
          localStorage.getItem(PENDING_UPDATE_KEY);

        if (pendingRaw) {
          try {
            const pending =
              JSON.parse(pendingRaw) as UpdateNotice;

            if (pending.version === currentVersion) {
              localStorage.removeItem(
                PENDING_UPDATE_KEY,
              );

              if (!cancelled) {
                setInstalledNotice(pending);
                await revealMainWindow();
              }

              return;
            }
          } catch {
            localStorage.removeItem(
              PENDING_UPDATE_KEY,
            );
          }
        }

        const update = await check({
          timeout: 15000,
        });

        if (cancelled) {
          await update?.close();
          return;
        }

        if (update) {
          setAvailableUpdate(update);
          setUpdateOpen(true);
          await revealMainWindow();
        }
      } catch (error) {
        console.error(
          "Could not check for application updates:",
          error,
        );
      }
    }

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, []);

  async function downloadUpdate() {
    if (
      !availableUpdate ||
      downloading ||
      downloaded
    ) {
      return;
    }

    setDownloading(true);
    setDownloadProgress(0);
    setUpdateError("");

    let downloadedBytes = 0;
    let totalBytes = 0;

    try {
      await availableUpdate.download((event) => {
        if (event.event === "Started") {
          totalBytes =
            event.data.contentLength ?? 0;
          return;
        }

        if (event.event === "Progress") {
          downloadedBytes +=
            event.data.chunkLength;

          if (totalBytes > 0) {
            setDownloadProgress(
              Math.min(
                100,
                Math.round(
                  (downloadedBytes / totalBytes) *
                    100,
                ),
              ),
            );
          }

          return;
        }

        if (event.event === "Finished") {
          setDownloadProgress(100);
        }
      });

      setDownloaded(true);
      setDownloadProgress(100);
    } catch (error) {
      setUpdateError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    } finally {
      setDownloading(false);
    }
  }

  async function installUpdate() {
    if (
      !availableUpdate ||
      !downloaded ||
      installing
    ) {
      return;
    }

    setInstalling(true);
    setUpdateError("");

    localStorage.setItem(
      PENDING_UPDATE_KEY,
      JSON.stringify({
        version: availableUpdate.version,
        notes:
          availableUpdate.body?.trim() ||
          "General improvements and fixes.",
      }),
    );

    try {
      await availableUpdate.install();
    } catch (error) {
      localStorage.removeItem(
        PENDING_UPDATE_KEY,
      );

      setUpdateError(
        error instanceof Error
          ? error.message
          : String(error),
      );

      setInstalling(false);
    }
  }

  async function dismissUpdate() {
    if (downloading || installing) {
      return;
    }

    const update = availableUpdate;

    setUpdateOpen(false);
    setAvailableUpdate(null);
    setDownloaded(false);
    setDownloadProgress(null);
    setUpdateError("");

    try {
      await update?.close();
    } catch {
      // The updater resource may already be closed.
    }
  }

  if (installedNotice) {
    return (
      <div
        className="catalogue-update-overlay"
        style={{ zIndex: 10000 }}
      >
        <section className="catalogue-update-modal">
          <div className="catalogue-update-heading">
            <span className="catalogue-update-kicker">
              UPDATE INSTALLED
            </span>

            <h2>
              Updated to PoE 2 Unique Tracker{" "}
              {installedNotice.version}
            </h2>

            <p>
              PoE 2 Unique Tracker was updated successfully.
            </p>
          </div>

          <p
            className="catalogue-update-note"
            style={{ whiteSpace: "pre-wrap" }}
          >
            {installedNotice.notes}
          </p>

          <div className="catalogue-update-actions">
            <button
              type="button"
              className="catalogue-update-primary"
              onClick={() =>
                setInstalledNotice(null)
              }
            >
              Continue
            </button>
          </div>
        </section>
      </div>
    );
  }

  if (!updateOpen || !availableUpdate) {
    return null;
  }

  const notes =
    availableUpdate.body?.trim() ||
    "This update contains improvements and fixes.";

  return (
    <div
      className="catalogue-update-overlay"
      style={{ zIndex: 10000 }}
    >
      <section className="catalogue-update-modal">
        <div className="catalogue-update-heading">
          <span className="catalogue-update-kicker">
            {downloaded
              ? "UPDATE READY"
              : "APP UPDATE AVAILABLE"}
          </span>

          <h2>
            PoE 2 Unique Tracker {availableUpdate.version}
          </h2>

          <p>
            {downloaded
              ? "The update has finished downloading. PoE 2 Unique Tracker will close while Windows installs it."
              : `You are currently using version ${availableUpdate.currentVersion}.`}
          </p>
        </div>

        <p
          className="catalogue-update-note"
          style={{ whiteSpace: "pre-wrap" }}
        >
          <strong>What’s new</strong>
          <br />
          {notes}
        </p>

        {(downloading || downloaded) && (
          <>
            <div
              style={{
                height: 8,
                margin: "18px 26px 0",
                overflow: "hidden",
                borderRadius: 999,
                background: "#332e28",
              }}
            >
              <div
                style={{
                  width: `${downloadProgress ?? 0}%`,
                  height: "100%",
                  background: "#c69b59",
                  transition: "width 150ms ease",
                }}
              />
            </div>

            <p className="catalogue-update-note">
              {downloaded
                ? "Download complete."
                : downloadProgress === null
                  ? "Downloading update..."
                  : `Downloading update… ${downloadProgress}%`}
            </p>
          </>
        )}

        {updateError && (
          <p
            className="catalogue-update-note"
            style={{ color: "#d8897b" }}
          >
            Update failed: {updateError}
          </p>
        )}

        <div className="catalogue-update-actions">
          <button
            type="button"
            className="catalogue-update-secondary"
            disabled={downloading || installing}
            onClick={() => void dismissUpdate()}
          >
            Later
          </button>

          <button
            type="button"
            className="catalogue-update-primary"
            disabled={downloading || installing}
            onClick={() =>
              void (
                downloaded
                  ? installUpdate()
                  : downloadUpdate()
              )
            }
          >
            {installing
              ? "Starting Installer..."
              : downloading
                ? "Downloading..."
                : downloaded
                  ? "Install Update"
                  : "Update"}
          </button>
        </div>
      </section>
    </div>
  );
}