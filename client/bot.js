require("dotenv").config();
const { Client, GatewayIntentBits, Partials } = require('discord.js');

const client = new Client({
    intents: [GatewayIntentBits.Guilds], partials: [Partials.Channel]
});

client.on("ready", () => {
    console.log("Bot launched");
});

client.login(process.env.TOKEN);
