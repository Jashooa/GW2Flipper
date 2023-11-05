# GW2Flipper

## Overview
GW2Flipper is an educational exploration into non-invasive automation techniques using the popular online game, Guild Wars 2. It's designed to help you identify profitable in-game items and automate the process of buying and selling them to generate profits.

This project serves as a learning tool for understanding automation, data analysis, and financial strategies in the context of Guild Wars 2.

## Features
- Identify profitable in-game items: GW2Flipper uses https://www.gw2bltc.com/ to identify items that have the potential for profit, this list is then refined based on configurable parameters.
- Automated trading: Once profitable items are identified, the application can automatically buy and sell these items on the in-game trading post.
- Educational: GW2Flipper is designed to teach users about automation, data analysis, and market strategies in a non-invasive manner.

## Usage
To get started with GW2Flipper, follow these steps:

1. Clone the repository to your local machine.
2. Install the required dependencies.
3. Compile.
4. Configure your Guild Wars 2 API key and character name in config.json then copy it over to compile directory.
5. If needed compile and run BlacklistGenerator to generate a new blacklist.json containing a list of items that are incompatible with this application.
6. Run application.

## Configuration
Included is a sample config.json:
```json
{
  "ApiKey": "", // Add your Guild Wars 2 API key here, get one from https://account.arena.net/applications
  "CharacterName": "", // Add your character name here
  "BuysPerSellLoop": 50, // How many buy orders to make before doing a sell loop
  "Quantity": 0.01, // Multiplier of daily sells to buy
  "MaxSpend": 100000, // Max amount to spend on a buy order in copper
  "UndercutBuys": true, // Undercut buy orders
  "UndercutSells": true, // Undercut sell orders
  "BuyIfSelling": true, // Keep buying if you have a current sell order up
  "SellForBestIfUnderRange": true, // If lowest price is below acceptable profit then sell for profitable price
  "RemoveUnprofitable": true, // Remove items from buy list if they become unprofitable
  "IgnoreSellsLessThanBuys": true, // Ignore items with less daily sells than daily buys
  "SellsLessThanBuysRange": 1.5, // Acceptable range for daily sells to go over daily buys 
  "ErrorRange": 0.8, // Acceptable range of buy and sell price
  "ProfitRange": 0.5, // Acceptable range for profit
  "UpdateListTime": 15, // How often to update buy list in minutes
  "Arguments": [ // Website arguments for generating buy list from gw2bltc.com
    {
      "sell-min": "1",
      "buy-min": "1",
      "profit-min": "10",
      "profit-max": "100",
      "profit-pct-min": "15",
      "profit-pct-max": "60",
      "demand-min": "1000",
      "sold-day-min": "1000",
      "bought-day-min": "200"
    }
  ]
}

```

Also included is ocrfixes.json where common OCR errors can be adjusted.
