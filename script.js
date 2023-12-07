// Récupérez tous les conteneurs dans un tableau
const containers = ['loginContainer', 'contactContainer', 'featuresContainer', 'aboutUsContainer', 'aboutMeContainer'];

// Associez chaque bouton à son conteneur respectif
const buttonContainerMap = {
    'toggleLoginFormButton': 'loginContainer',
    'toggleContactButton': 'contactContainer',
    'toggleFeaturesButton': 'featuresContainer',
    'toggleAboutUsButton': 'aboutUsContainer',
    'toggleAboutMeButton': 'aboutMeContainer'
};

// Ajoutez un écouteur d'événements à chaque bouton
for (const buttonId in buttonContainerMap) {
    if (buttonContainerMap.hasOwnProperty(buttonId)) {
        const button = document.getElementById(buttonId);
        button.addEventListener('click', function () {
            const targetContainer = buttonContainerMap[buttonId];

            // Afficher ou cacher le conteneur en fonction de son état actuel
            const targetElement = document.getElementById(targetContainer);
            targetElement.style.display = targetElement.style.display === 'none' ? 'block' : 'none';

            // Masquer tous les autres conteneurs
            containers.filter(container => container !== targetContainer).forEach(container => {
                const element = document.getElementById(container);
                element.style.display = 'none';
            });
        });
    }
}

// Ajoutez un écouteur d'événements pour masquer tous les contenus en dehors des boutons
document.addEventListener('click', function (event) {
    const clickedElement = event.target;

    // Vérifiez si l'élément cliqué n'est pas un bouton
    if (!clickedElement.matches('button')) {
        // Masquer tous les conteneurs
        containers.forEach(container => {
            const element = document.getElementById(container);
            element.style.display = 'none';
        });
    }
});

function login(event) {
    event.preventDefault();
    var username = document.getElementById("username").value;
    var password = document.getElementById("password").value;
    console.log("Username: " + username);
    console.log("Password: " + password);
}
