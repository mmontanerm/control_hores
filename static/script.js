document.addEventListener("DOMContentLoaded", () => {
    const boto = document.getElementById("marcar-btn");
    const formulari = document.getElementById("formulari");
    const confirmar = document.getElementById("confirmar-btn");

    const ofSelect = document.getElementById("of-select");
    const ofAltres = document.getElementById("of-altre");

    const llocSelect = document.getElementById("lloc-select");
    const llocAltres = document.getElementById("lloc-altre");

    const comentaris = document.getElementById("comentaris");

    let mode = estatInicial === "ENTRADA" ? "ENTRADA" : "SORTIDA";

    const actualitzaBoto = () => {
        boto.textContent = "MARCAR " + mode;
        boto.className = mode === "ENTRADA" ? "entrada" : "sortida";
    };

    actualitzaBoto();

    boto.addEventListener("click", () => {
        if (mode === "ENTRADA") {
            formulari.style.display = "block";
            boto.style.display = "none";
        } else {
            fetch("/marcar", {
                method: "POST",
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ tipus: "SORTIDA" })
            }).then(() => {
                mode = "ENTRADA";
                actualitzaBoto();
            });
        }
    });

    confirmar.addEventListener("click", () => {
        const ofVal = ofSelect.value === "altre" ? ofAltres.value : ofSelect.value;
        const llocVal = llocSelect.value === "altre" ? llocAltres.value : llocSelect.value;

        fetch("/marcar", {
            method: "POST",
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                tipus: "ENTRADA",
                of: ofVal,
                lloc: llocVal,
                comentaris: comentaris.value
            })
        }).then(() => {
            formulari.style.display = "none";
            boto.style.display = "inline-block";
            mode = "SORTIDA";
            actualitzaBoto();
        });
    });

    ofSelect.addEventListener("change", () => {
        ofAltres.style.display = ofSelect.value === "altre" ? "inline-block" : "none";
    });

    llocSelect.addEventListener("change", () => {
        llocAltres.style.display = llocSelect.value === "altre" ? "inline-block" : "none";
    });
});
