/**
 * Utils.js - Implémentation minimale compatible avec les scripts existants
 * Ce fichier ajoute simplement la compatibilité avec la référence Utils
 * en redirigeant vers les fonctions globales existantes.
 */

// Création d'un objet Utils minimal qui délègue aux fonctions globales existantes
var Utils = {
  // Fonction idx - redirection vers la fonction globale idx
  idx: function(headerArray, name, def = -1) {
    return idx(headerArray, name, def);
  },
  
  // Fonction logAction - redirection vers la fonction globale logAction
  logAction: function(action) {
    logAction(action);
  }
};

// Pas besoin d'exposer quoi que ce soit d'autre, l'objet Utils est maintenant disponible globalement