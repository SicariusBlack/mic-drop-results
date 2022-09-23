require("dotenv").config();
const { Client, GatewayIntentBits, Partials } = require('discord.js');

const client = new Client({
    intents: [GatewayIntentBits.Guilds], partials: [Partials.Channel]
});

console.log(process.env.TOKEN)

client.on("ready", () => {
    console.log("Bot launched");
});

client.on("messageCreate", (message) => {
    if (message.content.startsWith("ping")) {
      message.channel.send("pong!");
    }
});

client.login(process.env.TOKEN);
