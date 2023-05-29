# Follow Up Action Logger

A Python script to pick up follow-up canvasser emails from your inbox and put them into
a Google Sheet for logging. 

## Installation

* Authorise API access via OAuth2 for Gmail and Google Sheets
* Set up an OAuth2 identity and put the `credentials.json` file into the same folder as
  this script
* If you want to use this script on a machine without a web browser, run it once on a
  machine with a web browser then copy the resulting `token.json` file across to your
  headless machine; it should then run without issue on the headless machine

## Usage

```
python follow-up-action-logger.py

```

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/#)
