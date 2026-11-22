# HTTP VB5

## Overview

The project is a library for handling HTTP requests and responses in VB5.

It provides classes and methods to create, send, and receive HTTP requests, as well as to parse and manipulate HTTP responses.

The library aims to enable HTTP communication in VB5 applications by offering a consistent and easy-to-use interface for common HTTP operations.

The library is designed to be compatible with VB5's limitations and conventions, ensuring that developers can integrate HTTP functionality without needing to switch to a more modern environment. It is designed to run on Windows NT4 with Internet Explorer 4.

This project starts with sockets and builds the HTTP functionality on top of them, providing a low-level foundation for network communication in VB5.

## Requirements

- Windows NT4
- Internet Explorer 4
- VB5 development environment
- Basic understanding of HTTP protocol and socket programming

# Rules

- Follow VB5 coding conventions and best practices.
- Ensure compatibility with Windows NT4 and Internet Explorer 4.
- Use sockets as the foundation for all HTTP communication.
- Handle HTTP requests and responses according to the HTTP/1.1 specification.
- Provide clear and consistent error handling for network and protocol errors.
- Document all classes, methods, and properties with usage examples.
- Avoid using features not supported by VB5 or the target environment.
- Write modular and reusable code to facilitate maintenance and extension.
- All code goes in the `src` directory.
- Do not run tests, and this project must be transferred as source code only.