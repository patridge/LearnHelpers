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
    const findModuleOrPathByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] In ('Module','Learning Path') AND [Custom.UID] Contains '{{uid}}'";
    const createUidModuleOrPathQuery = (uid) => {
        const query = findModuleOrPathByUidQuery.replace("{{uid}}", uid);
        return createAdHocQueryUrl(query);
    };
    // NOTE: Won't find issues when the module UID isn't present in the unit UIDs (e.g., differ in prefix).
    const findCustomerIssuesByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Customer Feedback' AND [Feedback Source]='Report an issue' AND [Custom.UID] Contains '{{uid}}' AND [State] NOT IN ('Closed', 'Rejected', 'Declined')";
    const createUidIssuesQuery = (uid) => {
        const query = findCustomerIssuesByUidQuery.replace("{{uid}}", uid);
        return createAdHocQueryUrl(query);
    };
    // NOTE: Won't find issues when the module UID isn't present in the unit UIDs (e.g., differ in prefix).
    const findVerbatimsByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Customer Feedback' AND [Feedback Source]='Star rating verbatim' AND [Custom.UID] Contains '{{uid}}' AND [State] NOT IN ('Closed', 'Rejected', 'Declined')";
    const createUidVerbatimsQuery = (uid) => {
        const query = findVerbatimsByUidQuery.replace("{{uid}}", uid);
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
        const uidModuleOrPathQueryUrlAnchor = uidResult.querySelector(".uid-query-content-work-item-url");
        const uidReviewQueryUrlAnchor = uidResult.querySelector(".uid-query-review-url");
        const uidIssuesQueryUrlAnchor = uidResult.querySelector(".uid-issues-url");
        const uidVerbatimsQueryUrlAnchor = uidResult.querySelector(".uid-verbatims-url");

        uidSpan.innerText = `"${uid}"`;
        const uidQueryModuleOrPathUrl = createUidModuleOrPathQuery(uid);
        uidModuleOrPathQueryUrlAnchor.setAttribute("href", uidQueryModuleOrPathUrl);
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
