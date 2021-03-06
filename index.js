let sections = [];
let pages = [];

const cards = [];
let cardIndex = 0;

const msalConfig = {
    auth: {
        clientId: "ba2f1a99-9e68-453e-a298-b21c86520911",
        authority: "https://login.microsoftonline.com/common"
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
    }
};

const msal = new Msal.UserAgentApplication(msalConfig);
const options = new MicrosoftGraph.MSALAuthenticationProviderOptions(["notes.read"]);
const authProvider = new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msal, options);

const Client = MicrosoftGraph.Client;
const client = Client.initWithMiddleware({ authProvider });

fetchNotebooks();

$("#logout").show();

async function fetchNotebooks() {

    const noteListPromise = await client.api("/me/onenote/notebooks").get();

    const notebooks = noteListPromise.value.map(n => { return { name: n.displayName, id: n.id } });
    
    html = "";

    notebooks.forEach(n => {

        html += `<button onclick="setNotebook('${n.id}')" class='list-group-item list-group-item-action'>${n.name}</button>`;

    });

    $("#loading").removeClass("d-flex");
    $("#loading").addClass("d-none");
    $("#choose-notebook").show();
    $("#notebooks").append(html);

}

function setNotebook(id) {

    $("#choose-notebook").hide();
    $("#loading").addClass("d-flex");
    $("#loading").removeClass("d-none");

    fetchNotebookData(id);

}

async function fetchNotebookData(notebook) {

    try {

        const sectPromise = client.api(`/me/onenote/notebooks/${notebook}/sections`).get();
        const pagePromise = client.api(`/me/onenote/pages`).filter(`parentNotebook/id eq '${notebook}'`).top(50).get();

        Promise.all([sectPromise, pagePromise]).then(res => {

            sections = res[0].value.map(n => { return {name: n.displayName, id: n.id} });
            pages = res[1].value.map(n => { return { name: n.title, id: n.id, sectId: n.parentSection.id } });

            console.log(sections, pages);

            setsEl = "";

            sections.forEach(n => {

                setsEl += `
                    <li class='list-group-item'>
                        <div class="custom-control custom-checkbox float-left">
                            <input type="checkbox" class="custom-control-input sect-title" id="${n.id}" onclick="selectAll('${n.id}')" />
                            <label class="custom-control-label" for="${n.id}"><h5>${n.name}</h5></label>
                        </div>
                    </li>
                `;

                subpages = pages.filter(x => x.sectId === n.id);
                subpages.forEach(x => setsEl += `
                    <li class='list-group-item'>
                        <div class="custom-control custom-checkbox float-left">
                            <input type="checkbox" class="custom-control-input ${x.sectId}" id="${x.id}" />
                            <label class="custom-control-label" for="${x.id}" style="margin-left:2em">${x.name}</label>
                        </div>
                    </li>
                `);

            });

            $("#loading").removeClass("d-flex");
            $("#loading").addClass("d-none");
            $("#sets").append(setsEl);
            $("#choose-cards").show();

        });

    }
    catch (err) { throw err };

}

async function card() {

    const batchReqArr = [];

    $("input[type='checkbox']").not(".sect-title").filter(":checked").each((n, el) => {
        
        batchReqArr.push({
            id: n,
            request: new Request(`/me/onenote/pages/${el.id}/content`),
            method: "GET"
        });  

    });

    const batchReq = new MicrosoftGraph.BatchRequestContent(batchReqArr);
    const content = await batchReq.getContent();

    const res = await client.api("/$batch").post(content);
    const batchRes = new MicrosoftGraph.BatchResponseContent(res);

    let iterator = batchRes.getResponsesIterator();
    let data = iterator.next();

    const dom = new DOMParser();

    while (!data.done) {
        const pageData = atob(await data.value[1].text());
        const doc = dom.parseFromString(pageData, "text/html");
        const table = doc.querySelector("tbody");

        console.log(table);

        if (!table) {
            data = iterator.next();
            continue;
        }

        [...table.children].forEach(n => {

            console.log(n.children[0]);

            const term = n.children[0].textContent;

            console.log("Term: " + term);
            console.log([...n.children[1].children]);
            console.log(n.children[1].children);

            if (!n.children[1].children.length && n.children[1].textContent)
                cards.push({ term: term, clue: n.children[1].textContent });

            [...n.children[1].children].forEach(x => {
                if (x.textContent)
                    cards.push({ term: term, clue: x.textContent });
            });
        });

        data = iterator.next();
    }

    console.log(cards);

    $("#choose-cards").hide();
    $("#card").show();
    $("#card-control").show();

    loadCard();

}

function loadCard() {
    $("#term-header").hide();

    $("#term").text(cards[cardIndex].term);
    $("#clue").text(cards[cardIndex].clue);
}

function revealCard() {
    $("#term-header").show();
}

function prevCard() {

    cardIndex--;
    if (cardIndex < 0) cardIndex = cards.length - 1;

    loadCard();

}

function nextCard() {

    cardIndex++;
    if (cardIndex === cards.length) cardIndex = 0;

    loadCard();

}

function shuffle() {

    for (let i = cards.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [cards[i], cards[j]] = [cards[j], cards[i]];
    }

    loadCard();

}

function selectAll(id) {

    $(`.${id}`).prop("checked", !$(`.${id}`).prop("checked"));

}

function logout() {
    msal.logout();
}