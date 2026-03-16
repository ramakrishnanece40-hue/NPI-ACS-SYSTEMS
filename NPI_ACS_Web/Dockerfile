FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY NPI_ACS_Web.csproj .
RUN dotnet restore NPI_ACS_Web.csproj

COPY . .
RUN dotnet publish NPI_ACS_Web.csproj -c Release -o /app

FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app .

ENTRYPOINT ["dotnet", "NPI_ACS_Web.dll"]