(() => {
  "use strict";

  // Later vervangen door je online backend-URL.
  // Voor lokale tests:
  const BACKEND = "http://localhost:8000";

  const insideTrimble = window.self !== window.top;

  const $ = (id) => document.getElementById(id);

  let workspaceApi = null;
  let explorerApi = null;

  let project = null;
  let accessToken = null;

  let companionWindow = null;

  let selectedIfc = null;
  let selectedPdf = null;

  let mode = "trimble";
  let analysis = null;

  if ($("excel-link")) {
    $("excel-link").href = `${BACKEND}/api/audit.xlsx`;
  }

  // ------------------------------------------------------------
  // START
  // ------------------------------------------------------------

  if (insideTrimble) {
    startExtension();
  } else {
    startCompanion();
  }

  // ============================================================
  // TRIMBLE EXTENSION — SCHERM 1
  // ============================================================

  async function startExtension() {
    $("extension-ui").hidden = false;
    $("standalone-ui").hidden = true;

    $("status").textContent = "Verbinden met Trimble Connect…";

    try {
      workspaceApi = await TrimbleConnectWorkspace.connect(
        window.parent,
        onWorkspaceEvent,
        30000
      );

      // Huidig Trimble project
      project = await workspaceApi.project.getProject();

      // Vraag Trimble access token
      try {
        const permission =
          await workspaceApi.extension.requestPermission("accesstoken");

        if (
          typeof permission === "string" &&
          permission.length > 20
        ) {
          accessToken = permission;
        }
      } catch (error) {
        console.warn("Access token permission:", error);
      }

      $("status").textContent =
        `Verbonden: ${project?.name || project?.id || "Trimble Connect"}`;

      // Knop om scherm 2 te openen
      $("open-companion").addEventListener(
        "click",
        openCompanionWindow
      );

      // Berichten ontvangen vanuit scherm 2
      window.addEventListener(
        "message",
        onCompanionMessage
      );

    } catch (error) {
      console.error(error);

      $("status").textContent =
        `Trimble verbinding mislukt: ${error?.message || error}`;
    }
  }

  // ------------------------------------------------------------
  // Events die Trimble zelf naar de extension stuurt
  // ------------------------------------------------------------

  function onWorkspaceEvent(event, args) {
    const data = args?.data ?? args;

    // Access token kan ook via event binnenkomen
    if (event === "extension.accessToken") {
      if (
        typeof data === "string" &&
        data.length > 20
      ) {
        accessToken = data;

        sendContextToCompanion();
      }
    }

    if (event === "viewer.onSelectionChanged") {
      sendToCompanion({
        type: "viewer-selection",
        data
      });
    }
  }

  // ------------------------------------------------------------
  // Open scherm 2
  // ------------------------------------------------------------

  function openCompanionWindow() {
    const url = new URL(window.location.href);

    url.searchParams.set("companion", "1");

    companionWindow = window.open(
      url.toString(),
      "trimble-data-import-ai-control",
      "width=1450,height=950"
    );

    // Companion meldt normaal zelf wanneer hij klaar is.
    // Deze extra poging vangt trage loads op.
    setTimeout(() => {
      sendContextToCompanion();
    }, 1000);
  }

  // ------------------------------------------------------------
  // Bericht van scherm 2 naar Trimble extension
  // ------------------------------------------------------------

  async function onCompanionMessage(event) {
    // Alleen eigen GitHub Pages origin toestaan
    if (event.origin !== window.location.origin) {
      return;
    }

    const message = event.data;

    if (!message) return;

    // Companion zegt: ik ben geladen
    if (message.type === "companion-ready") {
      // event.source is het geopende tweede venster
      companionWindow = event.source;

      sendContextToCompanion();

      return;
    }

    // Companion vraagt om BIM-object te tonen
    if (
      message.type === "focus" &&
      message.guid
    ) {
      await focusGuid(
        message.guid,
        message.modelId
      );
    }
  }

  // ------------------------------------------------------------
  // Stuur project + token naar scherm 2
  // ------------------------------------------------------------

  function sendContextToCompanion() {
    sendToCompanion({
      type: "trimble-context",
      project,
      token: accessToken
    });
  }

  function sendToCompanion(message) {
    if (
      companionWindow &&
      !companionWindow.closed
    ) {
      companionWindow.postMessage(
        message,
        window.location.origin
      );
    }
  }

  // ------------------------------------------------------------
  // BIM object selecteren en naar toe zoomen
  // ------------------------------------------------------------

  async function focusGuid(
    guid,
    preferredModelId
  ) {
    if (!workspaceApi) return;

    try {
      const models =
        await workspaceApi.viewer.getModels();

      const preferred =
        preferredModelId
          ? models.filter(
              (model) =>
                model.id === preferredModelId
            )
          : [];

      const modelsToSearch =
        preferred.length
          ? preferred
          : models;

      for (const model of modelsToSearch) {

        const found =
          await workspaceApi.viewer.getObjects({
            modelObjectIds: [
              {
                modelId: model.id
              }
            ],
            parameter: {
              properties: {
                GlobalId: guid
              }
            }
          });

        if (!found?.length) {
          continue;
        }

        const selector = {
          modelObjectIds:
            found
              .map((group) => ({
                modelId: group.modelId,

                objectIds:
                  (group.objects || [])
                    .map(
                      (object) =>
                        object.id ||
                        object.objectId ||
                        object.externalId
                    )
                    .filter(Boolean)
              }))
              .filter(
                (group) =>
                  group.objectIds.length
              )
        };

        if (
          selector.modelObjectIds.length
        ) {
          await workspaceApi.viewer.setSelection(
            selector,
            "set"
          );

          try {
            await workspaceApi.viewer.setCamera(
              selector,
              {
                animationTime: 350
              }
            );
          } catch (error) {
            console.warn(
              "Camera zoom niet beschikbaar:",
              error
            );
          }

          return;
        }
      }

      console.warn(
        "GUID niet gevonden in viewer:",
        guid
      );

    } catch (error) {
      console.error(
        "focusGuid fout:",
        error
      );
    }
  }

  // ============================================================
  // COMPANION — SCHERM 2
  // ============================================================

  function startCompanion() {
    $("extension-ui").hidden = true;
    $("standalone-ui").hidden = false;

    $("status").textContent =
      "Controlevenster";

    $("tab-trimble").addEventListener(
      "click",
      () => setMode("trimble")
    );

    $("tab-local").addEventListener(
      "click",
      () => setMode("local")
    );

    $("analyze").addEventListener(
      "click",
      analyze
    );

    $("confirm-high").addEventListener(
      "click",
      confirmHighConfidence
    );

    $("apply").addEventListener(
      "click",
      applyConfirmed
    );

    $("local-ifc").addEventListener(
      "change",
      updateAnalyzeButton
    );

    $("local-pdf").addEventListener(
      "change",
      updateAnalyzeButton
    );

    // Ontvang project/token uit Trimble extension
    window.addEventListener(
      "message",
      onExtensionMessage
    );

    // Vertel opener dat scherm 2 klaar is
    if (window.opener) {
      window.opener.postMessage(
        {
          type: "companion-ready"
        },
        window.location.origin
      );
    }

    updateAnalyzeButton();
  }

  // ------------------------------------------------------------
  // Bericht uit Trimble extension ontvangen
  // ------------------------------------------------------------

  function onExtensionMessage(event) {
    if (
      event.origin !== window.location.origin
    ) {
      return;
    }

    const message = event.data;

    if (!message) return;

    if (
      message.type === "trimble-context"
    ) {
      project =
        message.project || project;

      accessToken =
        message.token || accessToken;

      if (
        project?.id &&
        accessToken
      ) {
        $("explorer-status").textContent =
          `Verbonden met project: ${
            project.name || project.id
          }`;

        initializeTrimbleExplorer();
      }
    }

    if (
      message.type ===
      "viewer-selection"
    ) {
      console.log(
        "Viewer selectie:",
        message.data
      );
    }
  }

  // ------------------------------------------------------------
  // Trimble / lokaal tabs
  // ------------------------------------------------------------

  function setMode(nextMode) {
    mode = nextMode;

    $("tab-trimble").classList.toggle(
      "active",
      mode === "trimble"
    );

    $("tab-local").classList.toggle(
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
      !accessToken ||
      explorerApi
    ) {
      return;
    }

    try {
      $("explorer-status").textContent =
        "Trimble File Explorer laden…";

      const iframe =
        $("trimble-explorer");

      iframe.src =
        TrimbleConnectWorkspace.getConnectEmbedUrl();

      explorerApi =
        await TrimbleConnectWorkspace.connect(
          iframe,
          onExplorerEvent,
          30000
        );

      await explorerApi.embed.setTokens({
        accessToken
      });

      await explorerApi.embed.initFileExplorer({
        projectId: project.id,

        enableSelect: true,

        enableUploadFiles: false,

        enableCreateFolder: false,

        enableExplorerKebabMenu: false,

        fileTypeFilter: [
          "ifc",
          "pdf"
        ]
      });

      $("explorer-status").textContent =
        `Project: ${
          project.name || project.id
        }`;

    } catch (error) {
      console.error(error);

      $("explorer-status").textContent =
        `Trimble File Explorer kon niet laden: ${
          error?.message || error
        }`;
    }
  }

  // ------------------------------------------------------------
  // File Explorer selectie
  // ------------------------------------------------------------

  function onExplorerEvent(event, args) {
    if (
      event !== "extension.fileSelected"
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

    const name =
      (file.name || "")
        .toLowerCase();

    if (
      name.endsWith(".ifc")
    ) {
      selectedIfc = file;

      $("ifc-name").textContent =
        file.name;
    }

    if (
      name.endsWith(".pdf")
    ) {
      selectedPdf = file;

      $("pdf-name").textContent =
        file.name;
    }

    updateAnalyzeButton();
  }

  // ============================================================
  // ANALYSE BUTTON STATE
  // ============================================================

  function updateAnalyzeButton() {
    let enabled = false;

    if (mode === "trimble") {
      enabled = Boolean(
        accessToken &&
        project?.id &&
        selectedIfc &&
        selectedPdf
      );
    } else {
      enabled = Boolean(
        $("local-ifc").files?.[0] &&
        $("local-pdf").files?.[0]
      );
    }

    $("analyze").disabled =
      !enabled;
  }

  // ============================================================
  // ANALYSE
  // ============================================================

  async function analyze() {
    setMessage("Analyseren…");

    $("analyze").disabled = true;

    try {
      let response;

      // --------------------------------
      // Vanuit Trimble bestanden
      // --------------------------------

      if (mode === "trimble") {
        response = await fetch(
          `${BACKEND}/api/analyze/trimble`,
          {
            method: "POST",

            headers: {
              "Content-Type":
                "application/json"
            },

            body: JSON.stringify({
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

      // --------------------------------
      // Lokale fallback
      // --------------------------------

      else {
        const form =
          new FormData();

        form.append(
          "ifc",
          $("local-ifc").files[0]
        );

        form.append(
          "pdf",
          $("local-pdf").files[0]
        );

        response = await fetch(
          `${BACKEND}/api/analyze/local`,
          {
            method: "POST",
            body: form
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

      renderResults();

      setMessage("");

    } catch (error) {
      setMessage(
        error?.message ||
        String(error)
      );

    } finally {
      updateAnalyzeButton();
    }
  }

  // ============================================================
  // RESULTS
  // ============================================================

  function renderResults() {
    $("results").hidden = false;

    $("count-elements").textContent =
      analysis.element_count ?? 0;

    $("count-assignments").textContent =
      analysis.assignment_count ?? 0;

    $("count-proposals").textContent =
      analysis.rows?.length ?? 0;

    const body =
      $("results-body");

    body.innerHTML = "";

    for (
      const [index, row]
      of (analysis.rows || []).entries()
    ) {

      const tr =
        document.createElement("tr");

      const confidenceClass =
        row.confidence >= 95
          ? "green"
          : row.confidence >= 75
            ? "amber"
            : "red";

      tr.innerHTML = `
        <td>
          <strong>
            ${escapeHtml(row.element_ref || "—")}
          </strong>

          <small>
            ${escapeHtml(row.guid)}
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
          <span class="conf ${confidenceClass}">
            ${Number(row.confidence || 0)}%
          </span>
        </td>

        <td>
          ${escapeHtml(row.method || "")}
        </td>

        <td>
          <select
            data-role="status"
            data-index="${index}"
          >
            <option value="proposed">
              Voorstel
            </option>

            <option value="confirmed">
              Bevestigd
            </option>

            <option value="needs_review">
              Controle nodig
            </option>

            <option value="rejected">
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

      body.appendChild(tr);

      const statusSelect =
        tr.querySelector(
          '[data-role="status"]'
        );

      statusSelect.value =
        row.status || "proposed";
    }

    // Waarde aanpassen
    body
      .querySelectorAll(
        '[data-role="value"]'
      )
      .forEach((element) => {
        element.addEventListener(
          "input",
          (event) => {
            const index =
              Number(
                event.target.dataset.index
              );

            analysis.rows[index].value =
              event.target.value;
          }
        );
      });

    // Status aanpassen
    body
      .querySelectorAll(
        '[data-role="status"]'
      )
      .forEach((element) => {
        element.addEventListener(
          "change",
          (event) => {
            const index =
              Number(
                event.target.dataset.index
              );

            analysis.rows[index].status =
              event.target.value;

            updateApplyButton();
          }
        );
      });

    // Bekijk in Trimble
    body
      .querySelectorAll(
        '[data-role="focus"]'
      )
      .forEach((element) => {

        element.addEventListener(
          "click",
          (event) => {

            const index =
              Number(
                event.target.dataset.index
              );

            const row =
              analysis.rows[index];

            if (window.opener) {
              window.opener.postMessage(
                {
                  type: "focus",
                  guid: row.guid,
                  modelId: row.model_id
                },
                window.location.origin
              );
            }
          }
        );
      });

    updateApplyButton();
  }

  // ============================================================
  // CONFIRM HIGH CONFIDENCE
  // ============================================================

  function confirmHighConfidence() {
    for (
      const row
      of (analysis?.rows || [])
    ) {
      if (
        Number(row.confidence) >= 95
      ) {
        row.status =
          "confirmed";
      }
    }

    renderResults();
  }

  function updateApplyButton() {
    const confirmed =
      (analysis?.rows || [])
        .some(
          (row) =>
            row.status === "confirmed"
        );

    $("apply").disabled =
      !confirmed;
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
        "Open dit controlevenster vanuit de Trimble extension."
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
            method: "POST",

            headers: {
              "Content-Type":
                "application/json"
            },

            body: JSON.stringify({
              access_token:
                accessToken,

              project_id:
                project.id,

              model_name:
                analysis.model_name,

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
          result.written || 0
        } verwerkt, ${
          result.failed || 0
        } fouten. Mode: ${
          result.mode || "onbekend"
        }`
      );

    } catch (error) {
      setMessage(
        error?.message ||
        String(error)
      );
    }
  }

  // ============================================================
  // HELPERS
  // ============================================================

  function setMessage(text) {
    const element =
      $("message");

    element.textContent =
      text || "";

    element.hidden =
      !text;
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;");
  }

  function escapeAttribute(value) {
    return escapeHtml(value)
      .replaceAll('"', "&quot;");
  }

})();
