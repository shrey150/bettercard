
window.addEventListener("keydown", e => {

    switch (e.key) {

        case " ":
        case "Enter":
            revealCard();
            break;

        case "n":
        case "d":
        case "ArrowRight":
            nextCard();
            break;

        case "p":
        case "a":
        case "ArrowLeft":
            prevCard();
            break;

    }

});