(function () {
    'use strict'

    const createAdHocQueryUrl = (query) => {
        return `https://ceapex.visualstudio.com/Microsoft%20Learn/_queries/query/?wiql=${encodeURIComponent(query)}`;
    };
    const findTechnicalReviewByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Technical Review' AND [Custom.UID] Contains '{{uid}}'";
    const createUidReviewQuery = (uid) => {
        const query = findTechnicalReviewByUidQuery.replace("{{uid}}", uid);
        return createAdHocQueryUrl(query);
    };
    const findModuleByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Module' AND [Custom.UID] Contains '{{uid}}'";
    const createUidModuleQuery = (uid) => {
        const query = findModuleByUidQuery.replace("{{uid}}", uid);
        return createAdHocQueryUrl(query);
    };
    // NOTE: Won't find issues when the module UID isn't present in the unit UIDs (e.g., differ in prefix).
    const findModuleIssuesByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Customer Feedback' AND [Feedback Source]='Report an issue' AND [Custom.UID] Contains '{{uid}}' AND [State] NOT IN ('Closed', 'Rejected', 'Declined')";
    const createUidIssuesQuery = (uid) => {
        const query = findModuleIssuesByUidQuery.replace("{{uid}}", uid);
        return createAdHocQueryUrl(query);
    };
    // NOTE: Won't find issues when the module UID isn't present in the unit UIDs (e.g., differ in prefix).
    const findModuleVerbatimsByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Customer Feedback' AND [Feedback Source]='Star rating verbatim' AND [Custom.UID] Contains '{{uid}}' AND [State] NOT IN ('Closed', 'Rejected', 'Declined')";
    const createUidVerbatimsQuery = (uid) => {
        const query = findModuleVerbatimsByUidQuery.replace("{{uid}}", uid);
        return createAdHocQueryUrl(query);
    };

    const forms = document.querySelectorAll(".needs-validation");
    const outputList = document.getElementById("uid-queries");
    const uidForm = document.getElementById("uid-query");
    const uidInput = document.getElementById("uid");
    const uidResultTemplate = document.getElementById("template-uid-result");

    const addUidResult = function (uid) {
        const uidResult = uidResultTemplate.content.cloneNode(true);
        const uidSpan = uidResult.querySelector(".uid");
        const uidModuleQueryUrlAnchor = uidResult.querySelector(".uid-query-module-url");
        const uidReviewQueryUrlAnchor = uidResult.querySelector(".uid-query-review-url");
        const uidIssuesQueryUrlAnchor = uidResult.querySelector(".uid-issues-url");
        const uidVerbatimsQueryUrlAnchor = uidResult.querySelector(".uid-verbatims-url");

        uidSpan.innerText = `"${uid}"`;
        const uidQueryModulesUrl = createUidModuleQuery(uid);
        uidModuleQueryUrlAnchor.setAttribute("href", uidQueryModulesUrl);
        const uidQueryReviewsUrl = createUidReviewQuery(uid);
        uidReviewQueryUrlAnchor.setAttribute("href", uidQueryReviewsUrl);
        const uidIssuesUrl = createUidIssuesQuery(uid);
        uidIssuesQueryUrlAnchor.setAttribute("href", uidIssuesUrl);
        const uidVerbatimsUrl = createUidVerbatimsQuery(uid);
        uidVerbatimsQueryUrlAnchor.setAttribute("href", uidVerbatimsUrl)
        outputList.insertBefore(uidResult, outputList.firstChild);
    };

    uidInput.addEventListener("click", (event) => {
        console.log(event);
        if (event.target.value !== "") {
            event.target.select();
        }
    }, false);

    [...forms].forEach(form => {
        form.addEventListener("submit", (event) => {
            if (!form.checkValidity()) {
                event.preventDefault();
                event.stopImmediatePropagation();
                event.stopPropagation();
            }
            
            form.classList.add("was-validated");
        }, false);
    });
    uidForm.addEventListener("submit", (event) => {
        const uid = uidInput.value;
        addUidResult(uid);
        event.preventDefault();
    }, false);
})();
