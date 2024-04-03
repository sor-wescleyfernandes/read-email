"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.initialReadEmails = void 0;
const emailService_1 = require("../services/emailService");
// Função para iniciar a leitura de emails
const initialReadEmails = () => {
    (0, emailService_1.lerEmails)();
};
exports.initialReadEmails = initialReadEmails;
