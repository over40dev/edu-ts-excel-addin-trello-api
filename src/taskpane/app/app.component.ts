import {Component} from "@angular/core";
const template = require("./app.component.html");
/* global console, Excel, require */

@Component({selector: "app-home", template: template})
export default class AppComponent {
  welcomeMessage = "Welcome";
  boards = [];
  cards = [];

  getBoards() {

    this.boards = [];

    const boards = [
      {
        "id": "1",
        "name": "My Sample Board",
        "closed": false
      }, {
        "id": "2",
        "name": "My Second Sample Board",
        "closed": false
      }
    ];

    this
      .boards
      .push(...boards);

  }

  getCards() {

    const cards = [
      {
        "name": "Card 1",
        "desc": "Description 1"
      }, {
        "name": "Card 2",
        "desc": "Description 2"
      }
    ];

    this
      .cards
      .push(...cards);

  }

  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
        */
       console.log('here');

        //  end my code

        const range = context
          .workbook
          .getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
}
