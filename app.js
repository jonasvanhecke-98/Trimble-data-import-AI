(() => {
  "use strict";

  // ============================================================
  // CONFIG
  // ============================================================

  // Voorlopig lokaal.
  // Later vervangen door je online HTTPS backend.
  const BACKEND = "http://localhost:8000";

  const insideTrimble = window.self !== window.top;
  const $ = (id) => document.getElementById(id);

  // ============================================================
  // STATE
  // ============================================================

  let workspaceApi = null;
  let explorerApi = null;

  let project = null;
  let accessToken = null;
  let permissionStatus = null;

  let companionWindow = null;

  let selectedIfc = null;
  let selectedPdf = null;

  let mode = "trimble";
  let analysis = null;

  let companionHandshakeTimer = null;

  // ============================================================
  // INIT
  // ============================================================

  if ($("excel-link")) {
    $("excel-link").href = `${BACKEND}/api/audit.xlsx`;
  }

  if (insideTrimble) {
    startExtension();
  } else {
    startCompanion();
  }

  // ============================================================
  // SCHERM 1
  // TRIMBLE EXTENSION
  // ============================================================

  async function startExtension() {
    $("extension-ui").hidden = false;
    $("standalone-ui").hidden = true;

    setExtensionStatus("Verbinden met Trimble Connect…");

    console.log("[Trimble AI] Extension starten");

    try {
      workspaceApi = await TrimbleConnectWorkspace.connect(
        window.parent,
        onWorkspaceEvent,
        30000
      );

      console.log("[Trimble AI] Workspace API verbonden");

      // --------------------------------------------------------
      // Project ophalen
      // --------------------------------------------------------

      try {
        project = await workspaceApi.project.getProject();

        console.log(
          "[Trimble AI] Project:",
          project
        );

        setExtensionStatus(
          `Verbonden met project: ${
            project?.name ||
            project?.id ||
            "onbekend"
          }`
        );

        // Project al doorgeven,
        // ook al hebben we nog geen token.
        sendContextToCompanion();

      } catch (error) {
        console.error(
          "[Trimble AI] Project ophalen mislukt:",
          error
        );

        setExtensionStatus(
          "Verbonden met Trimble, maar project kon niet worden opgehaald."
        );
      }

      // --------------------------------------------------------
      // Access token toestemming aanvragen
      // --------------------------------------------------------

      try {
        permissionStatus =
          await workspaceApi.extension.requestPermission(
            "accesstoken"
          );

        console.log(
          "[Trimble AI] requestPermission resultaat:",
          permissionStatus
        );

        handlePermissionResult(
          permissionStatus
        );

      } catch (error) {
        console.error(
          "[Trimble AI] requestPermission fout:",
          error
        );

        permissionStatus = "error";

        updateExtensionStatus();
      }

      // --------------------------------------------------------
      // Open companion
      // --------------------------------------------------------

      $("open-companion")
        .addEventListener(
          "click",
          openCompanionWindow
        );

      // --------------------------------------------------------
      // Berichten vanuit companion
      // --------------------------------------------------------

      window.addEventListener(
        "message",
        onCompanionMessage
      );

    } catch (error) {
      console.error(
        "[Trimble AI] Workspace connect fout:",
        error
      );

      setExtensionStatus(
        `Trimble verbinding mislukt: ${
          error?.message || error
        }`
      );
    }
  }

  // ============================================================
  // WORKSPACE EVENTS
  // ============================================================

  function onWorkspaceEvent(
    event,
    args
  ) {
    console.log(
      "[Trimble AI] Workspace event:",
      event,
      args
    );

    const data =
      args?.data ?? args;

    // ----------------------------------------------------------
    // Access token event
    // ----------------------------------------------------------

    if (
      event ===
      "extension.accessToken"
    ) {
      console.log(
        "[Trimble AI] Access token event:",
        data
      );

      handlePermissionResult(data);
    }

    // ----------------------------------------------------------
    // Viewer selection
    // ----------------------------------------------------------

    if (
      event ===
      "viewer.onSelectionChanged"
    ) {
      sendToCompanion({
        type: "viewer-selection",
        data
      });
    }
  }

  // ============================================================
  // PERMISSION / TOKEN
  // ============================================================

  function handlePermissionResult(
    value
  ) {
    if (!value) {
      updateExtensionStatus();
      sendContextToCompanion();
      return;
    }

    // ----------------------------------------------------------
    // Pending
    // ----------------------------------------------------------

    if (value === "pending") {
      permissionStatus = "pending";

      console.log(
        "[Trimble AI] Access token wacht op gebruikerstoestemming."
      );

      updateExtensionStatus();
      sendContextToCompanion();

      return;
    }

    // ----------------------------------------------------------
    // Denied
    // ----------------------------------------------------------

    if (value === "denied") {
      permissionStatus = "denied";

      console.warn(
        "[Trimble AI] Access token toestemming geweigerd."
      );

      updateExtensionStatus();
      sendContextToCompanion();

      return;
    }

    // ----------------------------------------------------------
    // Mogelijk object
    // ----------------------------------------------------------

    let possibleToken = value;

    if (
      typeof value === "object" &&
      value !== null
    ) {
      possibleToken =
        value.accessToken ||
        value.token ||
        value.data ||
        null;
    }

    // ----------------------------------------------------------
    // Token ontvangen
    // ----------------------------------------------------------

    if (
      typeof possibleToken ===
        "string" &&
      possibleToken.length > 20
    ) {
      accessToken =
        possibleToken;

      permissionStatus =
        "granted";

      console.log(
        "[Trimble AI] Access token ontvangen."
      );

      updateExtensionStatus();
      sendContextToCompanion();

      return;
    }

    console.log(
      "[Trimble AI] Onbekend permission resultaat:",
      value
    );

    permissionStatus =
      String(value);

    updateExtensionStatus();
    sendContextToCompanion();
  }

  function updateExtensionStatus() {
    const projectName =
      project?.name ||
      project?.id ||
      "Trimble project";

    if (accessToken) {
      setExtensionStatus(
        `Verbonden: ${projectName}`
      );

      return;
    }

    if (
      permissionStatus ===
      "pending"
    ) {
      setExtensionStatus(
        `Verbonden: ${projectName} — wacht op toestemming voor bestandstoegang`
      );

      return;
    }

    if (
      permissionStatus ===
      "denied"
    ) {
      setExtensionStatus(
        `Verbonden: ${projectName} — toegang tot bestanden geweigerd`
      );

      return;
    }

    if (
      permissionStatus ===
      "error"
    ) {
      setExtensionStatus(
        `Verbonden: ${projectName} — token kon niet worden aangevraagd`
      );

      return;
    }

    setExtensionStatus(
      `Verbonden: ${projectName}`
    );
  }

  function setExtensionStatus(
    text
  ) {
    if ($("status")) {
      $("status").textContent =
        text;
    }
  }

  // ============================================================
  // OPEN SCHERM 2
  // ============================================================

  function openCompanionWindow() {
    const url =
      new URL(
        window.location.href
      );

    url.searchParams.set(
      "companion",
      "1"
    );

    companionWindow =
      window.open(
        url.toString(),
        "trimble-data-import-ai-control",
        "width=1450,height=950"
      );

    if (!companionWindow) {
      alert(
        "Het controlevenster werd door de browser geblokkeerd. Sta pop-ups toe voor deze site."
      );

      return;
    }

    console.log(
      "[Trimble AI] Companion geopend."
    );

    // Meteen proberen
    sendContextToCompanion();

    // Nogmaals na laden
    setTimeout(
      sendContextToCompanion,
      500
    );

    setTimeout(
      sendContextToCompanion,
      1500
    );

    setTimeout(
      sendContextToCompanion,
      3000
    );
  }

  // ============================================================
  // BERICHTEN VAN SCHERM 2
  // ============================================================

  async function onCompanionMessage(
    event
  ) {
    if (
      event.origin !==
      window.location.origin
    ) {
      return;
    }

    const message =
      event.data;

    if (!message) return;

    console.log(
      "[Trimble AI] Bericht van companion:",
      message
    );

    // ----------------------------------------------------------
    // Companion klaar
    // ----------------------------------------------------------

    if (
      message.type ===
      "companion-ready"
    ) {
      companionWindow =
        event.source;

      sendContextToCompanion();

      return;
    }

    // ----------------------------------------------------------
    // Companion vraagt opnieuw context
    // ----------------------------------------------------------

    if (
      message.type ===
      "request-trimble-context"
    ) {
      companionWindow =
        event.source;

      sendContextToCompanion();

      return;
    }

    // ----------------------------------------------------------
    // Focus BIM object
    // ----------------------------------------------------------

    if (
      message.type === "focus" &&
      message.guid
    ) {
      await focusGuid(
        message.guid,
        message.modelId
      );

      return;
    }
  }

  // ============================================================
  // CONTEXT NAAR SCHERM 2
  // ============================================================

  function sendContextToCompanion() {
    const message = {
      type: "trimble-context",

      project,

      token:
        accessToken,

      permissionStatus
    };

    sendToCompanion(
      message
    );
  }

  function sendToCompanion(
    message
  ) {
    if (
      !companionWindow ||
      companionWindow.closed
    ) {
      return;
    }

    try {
      companionWindow.postMessage(
        message,
        window.location.origin
      );
    } catch (error) {
      console.error(
        "[Trimble AI] postMessage fout:",
        error
      );
    }
  }

  // ============================================================
  // VIEWER OBJECT SELECTEREN
  // ============================================================

  async function focusGuid(
    guid,
    preferredModelId
  ) {
    if (!workspaceApi) {
      console.warn(
        "[Trimble AI] Geen workspaceApi."
      );

      return;
    }

    try {
      console.log(
        "[Trimble AI] Zoek GUID:",
        guid
      );

      const models =
        await workspaceApi.viewer.getModels();

      console.log(
        "[Trimble AI] Viewer modellen:",
        models
      );

      const preferred =
        preferredModelId
          ? models.filter(
              (model) =>
                model.id ===
                preferredModelId
            )
          : [];

      const modelsToSearch =
        preferred.length
          ? preferred
          : models;

      for (
        const model
        of modelsToSearch
      ) {
        try {
          const found =
            await workspaceApi.viewer.getObjects({
              modelObjectIds: [
                {
                  modelId:
                    model.id
                }
              ],

              parameter: {
                properties: {
                  GlobalId:
                    guid
                }
              }
            });

          console.log(
            "[Trimble AI] GUID resultaat:",
            model.id,
            found
          );

          if (!found?.length) {
            continue;
          }

          const selector = {
            modelObjectIds:
              found
                .map(
                  (group) => ({
                    modelId:
                      group.modelId,

                    objectIds:
                      (
                        group.objects ||
                        []
                      )
                        .map(
                          (object) =>
                            object.id ||
                            object.objectId ||
                            object.externalId
                        )
                        .filter(
                          Boolean
                        )
                  })
                )
                .filter(
                  (group) =>
                    group.objectIds
                      .length
                )
          };

          if (
            selector
              .modelObjectIds
              .length
          ) {
            await workspaceApi
              .viewer
              .setSelection(
                selector,
                "set"
              );

            try {
              await workspaceApi
                .viewer
                .setCamera(
                  selector,
                  {
                    animationTime:
                      350
                  }
                );

            } catch (
              cameraError
            ) {
              console.warn(
                "[Trimble AI] Camera zoom fout:",
                cameraError
              );
            }

            return;
          }

        } catch (
          modelError
        ) {
          console.warn(
            "[Trimble AI] Zoekfout in model:",
            model.id,
            modelError
          );
        }
      }

      console.warn(
        "[Trimble AI] GUID niet gevonden:",
        guid
      );

    } catch (error) {
      console.error(
        "[Trimble AI] focusGuid fout:",
        error
      );
    }
  }

  // ============================================================
  // SCHERM 2 / COMPANION
  // ============================================================

  function startCompanion() {
    $("extension-ui").hidden =
      true;

    $("standalone-ui").hidden =
      false;

    $("status").textContent =
      "Controlevenster";

    // ----------------------------------------------------------
    // UI events
    // ----------------------------------------------------------

    $("tab-trimble")
      .addEventListener(
        "click",
        () =>
          setMode("trimble")
      );

    $("tab-local")
      .addEventListener(
        "click",
        () =>
          setMode("local")
      );

    $("analyze")
      .addEventListener(
        "click",
        analyze
      );

    $("confirm-high")
      .addEventListener(
        "click",
        confirmHighConfidence
      );

    $("apply")
      .addEventListener(
        "click",
        applyConfirmed
      );

    $("local-ifc")
      .addEventListener(
        "change",
        updateAnalyzeButton
      );

    $("local-pdf")
      .addEventListener(
        "change",
        updateAnalyzeButton
      );

    // ----------------------------------------------------------
    // Luister naar extension
    // ----------------------------------------------------------

    window.addEventListener(
      "message",
      onExtensionMessage
    );

    // ----------------------------------------------------------
    // Companion meldt zichzelf
    // ----------------------------------------------------------

    notifyExtensionCompanionReady();

    // Blijf context vragen tot
    // project + token beschikbaar zijn.
    companionHandshakeTimer =
      setInterval(
        () => {
          if (
            project?.id &&
            accessToken
          ) {
            clearInterval(
              companionHandshakeTimer
            );

            companionHandshakeTimer =
              null;

            return;
          }

          requestTrimbleContext();

        },
        1500
      );

    updateAnalyzeButton();
  }

  // ============================================================
  // COMPANION HANDSHAKE
  // ============================================================

  function notifyExtensionCompanionReady() {
    if (!window.opener) {
      $("explorer-status")
        .textContent =
        "Dit venster werd niet vanuit Trimble Connect geopend.";

      return;
    }

    try {
      window.opener.postMessage(
        {
          type:
            "companion-ready"
        },
        window.location.origin
      );

      console.log(
        "[Trimble AI] companion-ready gestuurd."
      );

    } catch (error) {
      console.error(
        "[Trimble AI] companion-ready fout:",
        error
      );
    }
  }

  function requestTrimbleContext() {
    if (!window.opener) {
      return;
    }

    try {
      window.opener.postMessage(
        {
          type:
            "request-trimble-context"
        },
        window.location.origin
      );

      console.log(
        "[Trimble AI] Trimble context opnieuw gevraagd."
      );

    } catch (error) {
      console.error(
        "[Trimble AI] Context request fout:",
        error
      );
    }
  }

  // ============================================================
  // BERICHTEN VAN EXTENSION
  // ============================================================

  function onExtensionMessage(
    event
  ) {
    if (
      event.origin !==
      window.location.origin
    ) {
      return;
    }

    const message =
      event.data;

    if (!message) return;

    console.log(
      "[Trimble AI] Bericht vanuit Trimble:",
      message
    );

    if (
      message.type !==
      "trimble-context"
    ) {
      return;
    }

    if (message.project) {
      project =
        message.project;
    }

    if (message.token) {
      accessToken =
        message.token;
    }

    permissionStatus =
      message.permissionStatus ||
      permissionStatus;

    updateCompanionConnectionStatus();
  }

  // ============================================================
  // COMPANION STATUS
  // ============================================================

  function updateCompanionConnectionStatus() {
    const projectName =
      project?.name ||
      project?.id ||
      null;

    // Geen project
    if (!project?.id) {
      $("explorer-status")
        .textContent =
        "Wachten op Trimble project…";

      updateAnalyzeButton();
      return;
    }

    // Project wel,
    // token nog niet.
    if (!accessToken) {
      if (
        permissionStatus ===
        "pending"
      ) {
        $("explorer-status")
          .textContent =
          `Project verbonden: ${projectName} — wacht op toestemming voor bestandstoegang`;

      } else if (
        permissionStatus ===
        "denied"
      ) {
        $("explorer-status")
          .textContent =
          `Project verbonden: ${projectName} — toegang tot bestanden geweigerd`;

      } else {
        $("explorer-status")
          .textContent =
          `Project verbonden: ${projectName} — wacht op access token`;
      }

      updateAnalyzeButton();
      return;
    }

    // Alles aanwezig
    $("explorer-status")
      .textContent =
      `Verbonden met project: ${projectName}`;

    if (
      companionHandshakeTimer
    ) {
      clearInterval(
        companionHandshakeTimer
      );

      companionHandshakeTimer =
        null;
    }

    initializeTrimbleExplorer();
    updateAnalyzeButton();
  }

  // ============================================================
  // MODES
  // ============================================================

  function setMode(
    nextMode
  ) {
    mode =
      nextMode;

    $("tab-trimble")
      .classList
      .toggle(
        "active",
        mode === "trimble"
      );

    $("tab-local")
      .classList
      .toggle(
        "active",
        mode === "local"
      );

    $("trimble-pane").hidden =
      mode !== "trimble";

    $("local-pane").hidden =
      mode !== "local";

    updateAnalyzeButton();
  }

  // ============================================================
  // TRIMBLE FILE EXPLORER
  // ============================================================

  async function initializeTrimbleExplorer() {
    if (
      !project?.id ||
      !accessToken
    ) {
      return;
    }

    if (explorerApi) {
      return;
    }

    console.log(
      "[Trimble AI] File Explorer initialiseren."
    );

    try {
      $("explorer-status")
        .textContent =
        "Trimble File Explorer laden…";

      const iframe =
        $("trimble-explorer");

      iframe.src =
        TrimbleConnectWorkspace
          .getConnectEmbedUrl();

      explorerApi =
        await TrimbleConnectWorkspace.connect(
          iframe,
          onExplorerEvent,
          30000
        );

      console.log(
        "[Trimble AI] Embedded workspace verbonden."
      );

      await explorerApi
        .embed
        .setTokens({
          accessToken
        });

      console.log(
        "[Trimble AI] Token op File Explorer gezet."
      );

      await explorerApi
        .embed
        .initFileExplorer({
          projectId:
            project.id,

          enableSelect:
            true,

          enableUploadFiles:
            false,

          enableCreateFolder:
            false,

          enableExplorerKebabMenu:
            false,

          fileTypeFilter: [
            "ifc",
            "pdf"
          ]
        });

      $("explorer-status")
        .textContent =
        `Project: ${
          project.name ||
          project.id
        }`;

      console.log(
        "[Trimble AI] File Explorer klaar."
      );

    } catch (error) {
      explorerApi = null;

      console.error(
        "[Trimble AI] File Explorer fout:",
        error
      );

      $("explorer-status")
        .textContent =
        `Trimble File Explorer kon niet laden: ${
          error?.message ||
          error
        }`;
    }
  }

  // ============================================================
  // FILE SELECTIE
  // ============================================================

  function onExplorerEvent(
    event,
    args
  ) {
    console.log(
      "[Trimble AI] Explorer event:",
      event,
      args
    );

    if (
      event !==
      "extension.fileSelected"
    ) {
      return;
    }

    const file =
      args?.data?.file ||
      args?.file;

    if (
      !file ||
      file.type !== "FILE"
    ) {
      return;
    }

    const filename =
      (
        file.name ||
        ""
      ).toLowerCase();

    if (
      filename.endsWith(
        ".ifc"
      )
    ) {
      selectedIfc =
        file;

      $("ifc-name")
        .textContent =
        file.name;
    }

    if (
      filename.endsWith(
        ".pdf"
      )
    ) {
      selectedPdf =
        file;

      $("pdf-name")
        .textContent =
        file.name;
    }

    updateAnalyzeButton();
  }

  // ============================================================
  // ANALYSE BUTTON
  // ============================================================

  function updateAnalyzeButton() {
    let enabled =
      false;

    if (
      mode === "trimble"
    ) {
      enabled =
        Boolean(
          accessToken &&
          project?.id &&
          selectedIfc &&
          selectedPdf
        );

    } else {
      enabled =
        Boolean(
          $("local-ifc")
            .files?.[0] &&
          $("local-pdf")
            .files?.[0]
        );
    }

    $("analyze").disabled =
      !enabled;
  }

  // ============================================================
  // ANALYSE
  // ============================================================

  async function analyze() {
    setMessage(
      "Analyseren…"
    );

    $("analyze").disabled =
      true;

    try {
      let response;

      // --------------------------------------------------------
      // Trimble bestanden
      // --------------------------------------------------------

      if (
        mode === "trimble"
      ) {
        response =
          await fetch(
            `${BACKEND}/api/analyze/trimble`,
            {
              method: "POST",

              headers: {
                "Content-Type":
                  "application/json"
              },

              body:
                JSON.stringify({
                  access_token:
                    accessToken,

                  project_id:
                    project.id,

                  project_location:
                    project.location,

                  ifc_file:
                    selectedIfc,

                  pdf_file:
                    selectedPdf
                })
            }
          );

      }

      // --------------------------------------------------------
      // Lokale bestanden
      // --------------------------------------------------------

      else {
        const form =
          new FormData();

        form.append(
          "ifc",
          $("local-ifc")
            .files[0]
        );

        form.append(
          "pdf",
          $("local-pdf")
            .files[0]
        );

        response =
          await fetch(
            `${BACKEND}/api/analyze/local`,
            {
              method:
                "POST",

              body:
                form
            }
          );
      }

      if (!response.ok) {
        throw new Error(
          await response.text()
        );
      }

      analysis =
        await response.json();

      console.log(
        "[Trimble AI] Analyse:",
        analysis
      );

      renderResults();

      setMessage("");

    } catch (error) {
      console.error(
        "[Trimble AI] Analyse fout:",
        error
      );

      setMessage(
        error?.message ||
        String(error)
      );

    } finally {
      updateAnalyzeButton();
    }
  }

  // ============================================================
  // RESULTATEN
  // ============================================================

  function renderResults() {
    $("results").hidden =
      false;

    $("count-elements")
      .textContent =
      analysis
        .element_count ??
      0;

    $("count-assignments")
      .textContent =
      analysis
        .assignment_count ??
      0;

    $("count-proposals")
      .textContent =
      analysis
        .rows?.length ??
      0;

    const body =
      $("results-body");

    body.innerHTML = "";

    for (
      const [
        index,
        row
      ]
      of (
        analysis.rows ||
        []
      ).entries()
    ) {
      const tr =
        document.createElement(
          "tr"
        );

      const confidenceClass =
        row.confidence >= 95
          ? "green"
          : row.confidence >= 75
            ? "amber"
            : "red";

      tr.innerHTML = `
        <td>
          <strong>
            ${escapeHtml(
              row.element_ref ||
              "—"
            )}
          </strong>

          <small>
            ${escapeHtml(
              row.guid
            )}
          </small>
        </td>

        <td>
          <input
            data-role="value"
            data-index="${index}"
            value="${escapeAttribute(
              row.value || ""
            )}"
          >
        </td>

        <td>
          <span
            class="conf ${confidenceClass}"
          >
            ${
              Number(
                row.confidence ||
                0
              )
            }%
          </span>
        </td>

        <td>
          ${escapeHtml(
            row.method ||
            ""
          )}
        </td>

        <td>
          <select
            data-role="status"
            data-index="${index}"
          >
            <option
              value="proposed"
            >
              Voorstel
            </option>

            <option
              value="confirmed"
            >
              Bevestigd
            </option>

            <option
              value="needs_review"
            >
              Controle nodig
            </option>

            <option
              value="rejected"
            >
              Afgewezen
            </option>
          </select>
        </td>

        <td>
          <button
            data-role="focus"
            data-index="${index}"
          >
            Bekijk
          </button>
        </td>
      `;

      body.appendChild(
        tr
      );

      const status =
        tr.querySelector(
          '[data-role="status"]'
        );

      status.value =
        row.status ||
        "proposed";
    }

    // ----------------------------------------------------------
    // Waarde wijzigen
    // ----------------------------------------------------------

    body
      .querySelectorAll(
        '[data-role="value"]'
      )
      .forEach(
        (element) => {

          element
            .addEventListener(
              "input",
              (event) => {

                const index =
                  Number(
                    event
                      .target
                      .dataset
                      .index
                  );

                analysis
                  .rows[index]
                  .value =
                  event
                    .target
                    .value;
              }
            );
        }
      );

    // ----------------------------------------------------------
    // Status wijzigen
    // ----------------------------------------------------------

    body
      .querySelectorAll(
        '[data-role="status"]'
      )
      .forEach(
        (element) => {

          element
            .addEventListener(
              "change",
              (event) => {

                const index =
                  Number(
                    event
                      .target
                      .dataset
                      .index
                  );

                analysis
                  .rows[index]
                  .status =
                  event
                    .target
                    .value;

                updateApplyButton();
              }
            );
        }
      );

    // ----------------------------------------------------------
    // Bekijk object in Trimble
    // ----------------------------------------------------------

    body
      .querySelectorAll(
        '[data-role="focus"]'
      )
      .forEach(
        (element) => {

          element
            .addEventListener(
              "click",
              (event) => {

                const index =
                  Number(
                    event
                      .target
                      .dataset
                      .index
                  );

                const row =
                  analysis
                    .rows[
                      index
                    ];

                if (
                  window.opener
                ) {
                  window.opener
                    .postMessage(
                      {
                        type:
                          "focus",

                        guid:
                          row.guid,

                        modelId:
                          row.model_id
                      },

                      window
                        .location
                        .origin
                    );
                }
              }
            );
        }
      );

    updateApplyButton();
  }

  // ============================================================
  // CONFIRM ≥95%
  // ============================================================

  function confirmHighConfidence() {
    for (
      const row
      of (
        analysis?.rows ||
        []
      )
    ) {
      if (
        Number(
          row.confidence
        ) >= 95
      ) {
        row.status =
          "confirmed";
      }
    }

    renderResults();
  }

  // ============================================================
  // APPLY BUTTON
  // ============================================================

  function updateApplyButton() {
    const hasConfirmed =
      (
        analysis?.rows ||
        []
      )
        .some(
          (row) =>
            row.status ===
            "confirmed"
        );

    $("apply").disabled =
      !hasConfirmed;
  }

  // ============================================================
  // APPLY
  // ============================================================

  async function applyConfirmed() {
    if (
      !accessToken ||
      !project?.id
    ) {
      setMessage(
        "Geen actieve Trimble-projectcontext beschikbaar."
      );

      return;
    }

    setMessage(
      "Bevestigde waarden verwerken…"
    );

    try {
      const response =
        await fetch(
          `${BACKEND}/api/apply`,
          {
            method:
              "POST",

            headers: {
              "Content-Type":
                "application/json"
            },

            body:
              JSON.stringify({
                access_token:
                  accessToken,

                project_id:
                  project.id,

                model_name:
                  analysis
                    .model_name,

                rows:
                  analysis.rows,

                confirmed_by:
                  "Trimble user"
              })
          }
        );

      if (!response.ok) {
        throw new Error(
          await response.text()
        );
      }

      const result =
        await response.json();

      setMessage(
        `Klaar: ${
          result.written ||
          0
        } verwerkt, ${
          result.failed ||
          0
        } fouten. Mode: ${
          result.mode ||
          "onbekend"
        }`
      );

    } catch (error) {
      console.error(
        "[Trimble AI] Apply fout:",
        error
      );

      setMessage(
        error?.message ||
        String(error)
      );
    }
  }

  // ============================================================
  // HELPERS
  // ============================================================

  function setMessage(
    text
  ) {
    const element =
      $("message");

    if (!element) return;

    element.textContent =
      text || "";

    element.hidden =
      !text;
  }

  function escapeHtml(
    value
  ) {
    return String(value)
      .replaceAll(
        "&",
        "&amp;"
      )
      .replaceAll(
        "<",
        "&lt;"
      )
      .replaceAll(
        ">",
        "&gt;"
      );
  }

  function escapeAttribute(
    value
  ) {
    return escapeHtml(
      value
    )
      .replaceAll(
        '"',
        "&quot;"
      );
  }

})();
