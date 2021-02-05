(function () {
    'use strict'

    const findTechnicalReviewByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Technical Review' AND [Custom.UID] Contains '{{uid}}'";
    const findModuleByUidQuery = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State] FROM workitems WHERE [System.TeamProject] = 'Microsoft Learn' AND [System.WorkItemType] = 'Module' AND [Custom.UID] Contains '{{uid}}'";

    const createUidReviewQuery = (uid) => {
        const query = findTechnicalReviewByUidQuery.replace("{{uid}}", uid);
        return `https://ceapex.visualstudio.com/Microsoft%20Learn/_queries/query/?wiql=${encodeURIComponent(query)}`;
    };
    const createUidModuleQuery = (uid) => {
        const query = findModuleByUidQuery.replace("{{uid}}", uid);
        return `https://ceapex.visualstudio.com/Microsoft%20Learn/_queries/query/?wiql=${encodeURIComponent(query)}`;
    };

    const forms = document.querySelectorAll(".needs-validation");
    const outputList = document.getElementById("uid-queries");
    const uidForm = document.getElementById("uid-query");
    const uidResultTemplate = document.getElementById("template-uid-result");
    const addUidResult = function (uid) {
        const uidResult = uidResultTemplate.content.cloneNode(true);
        const uidSpan = uidResult.querySelector(".uid");
        const uidModuleQueryUrlAnchor = uidResult.querySelector(".uid-query-module-url");
        const uidReviewQueryUrlAnchor = uidResult.querySelector(".uid-query-review-url");

        uidSpan.innerText = `"${uid}"`;
        const uidQueryModulesUrl = createUidModuleQuery(uid);
        uidModuleQueryUrlAnchor.setAttribute("href", uidQueryModulesUrl);
        const uidQueryReviewsUrl = createUidReviewQuery(uid);
        uidReviewQueryUrlAnchor.setAttribute("href", uidQueryReviewsUrl);
        outputList.appendChild(uidResult);
    };

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
        const uid = document.getElementById("uid").value;
        addUidResult(uid);
        event.preventDefault();
    }, false);
})();
