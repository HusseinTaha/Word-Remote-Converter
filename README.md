# Word-Remote-Converter

This project uses the remote services to host a converter from word files to pdf using interop.

## Getting Started

Add a reference in your project to the exe file, then use the below code to call the remote word converter.
```
Just reference the exe file to your project. and add the method below:
WordRemoteConverter.RemoteConverter converter = (WordRemoteConverter.RemoteConverter)Activator.GetObject(typeof(WordRemoteConverter.RemoteConverter),
                    "http://localhost:8989/RemoteConverter");
```

### Prerequisites

You need to have office installed and activated for the converter to work perfectly.


## Running the tests

Run the exectuable first, then run the web application. Upload a file and you can check that the file is converted.
 
## Deployment

You can deploy this solution by adding a windows service, this service will host the converter. After that you can hit the service by url + port and it will return the file path.

## Built With

* .Net 4.5 c#

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
