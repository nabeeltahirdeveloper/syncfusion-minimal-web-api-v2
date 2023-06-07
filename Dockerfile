FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS base
RUN apt-get update
RUN apt-get install -y libfreetype6
RUN apt-get install -y libfontconfig1 libicu-dev libharfbuzz0b
RUN apt-get update \
 && apt-get install --assume-yes apt-utils \
 && apt-get install --assume-yes software-properties-common \
 && apt-get install --assume-yes dbus \
 && apt-get install --assume-yes glib-networking \
 && apt-get install --assume-yes libnih-dbus-dev \
 && apt-get install --assume-yes dconf-cli \
 && apt-get install --assume-yes fontconfig

RUN rm -rf ./usr/share/fonts/truetype/
RUN rm -rf /usr/share/fonts/myfonts/

COPY ./static ./usr/share/fonts/truetype/
# Create a directory to store the fonts
RUN mkdir -p /usr/share/fonts/myfonts

# Copy the font files to the container
COPY static/* /usr/share/fonts/myfonts/
WORKDIR /app
EXPOSE 80
FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
WORKDIR /src
COPY . .
COPY ./static ./usr/share/fonts/truetype/
# Create a directory to store the fonts
RUN mkdir -p /usr/share/fonts/myfonts

# Copy the font files to the container
COPY static/* /usr/share/fonts/myfonts/
RUN dotnet --version

RUN dotnet restore
RUN dotnet build "DocumentEditorCore.csproj" -c Release -o /app/build
FROM build AS publish
RUN dotnet publish "DocumentEditorCore.csproj" -c Release -o /app/publish
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "DocumentEditorCore.dll"]