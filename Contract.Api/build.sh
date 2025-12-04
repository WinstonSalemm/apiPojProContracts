#!/bin/bash
dotnet publish -c Release -o out
cd out
./Contract.Api
