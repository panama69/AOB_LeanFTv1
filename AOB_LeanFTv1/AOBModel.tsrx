<?xml version="1.0" encoding="utf-8"?>
<applicationModel>
  <name>AOB Application Model</name>
  <className>AOBModel</className>
  <objects>
    <object technology="Web" type="Page">
      <name>Advantage Online Banking</name>
      <codeName>AdvantageOnlineBankingPage</codeName>
      <description>Main application login page</description>
      <identification />
      <children>
        <object technology="Web" type="EditField">
          <name>j_username</name>
          <codeName>UserName</codeName>
          <identification>
            <properties>
              <property name="Type" type="string">text</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">j_username</property>
            </properties>
          </identification>
          <tags>
            <tag>login</tag>
          </tags>
        </object>
        <object technology="Web" type="EditField">
          <name>j_password</name>
          <codeName>UserPassword</codeName>
          <identification>
            <properties>
              <property name="Type" type="string">password</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">j_password</property>
            </properties>
          </identification>
          <tags>
            <tag>login</tag>
          </tags>
        </object>
        <object technology="Web" type="Link">
          <name>Logout</name>
          <codeName>LogoutLink</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">A</property>
              <property name="InnerText" type="string">Logout</property>
            </properties>
          </identification>
          <tags>
            <tag>logout</tag>
            <tag>mainPage</tag>
          </tags>
        </object>
        <object technology="Web" type="Button">
          <name>Login</name>
          <codeName>LoginButton</codeName>
          <identification>
            <properties>
              <property name="ButtonType" type="string">submit</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">Login</property>
            </properties>
          </identification>
          <tags>
            <tag>login</tag>
          </tags>
        </object>
        <object technology="Web" type="Link">
          <name>Accounts </name>
          <codeName>AccountsLink</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">A</property>
              <property name="InnerText" type="string">Accounts </property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="Link">
          <name> Money Transfer </name>
          <codeName>MoneyTransferLink</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">A</property>
              <property name="InnerText" type="string"> Money Transfer </property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="ListBox">
          <name>fromAccount</name>
          <codeName>FromAccountListBox</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">SELECT</property>
              <property name="Name" type="string">fromAccount</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="ListBox">
          <name>toAccount</name>
          <codeName>ToAccountListBox</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">SELECT</property>
              <property name="Name" type="string">toAccount</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="Button">
          <name>Next</name>
          <codeName>NextButton</codeName>
          <identification>
            <properties>
              <property name="ButtonType" type="string">submit</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">Next</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="EditField">
          <name>amount</name>
          <codeName>AmountEditField</codeName>
          <identification>
            <properties>
              <property name="Type" type="string">text</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">amount</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="EditField">
          <name>transferDate</name>
          <codeName>TransferDateEditField</codeName>
          <identification>
            <properties>
              <property name="Type" type="string">text</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">transferDate</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="WebElement">
          <name>Reference Number</name>
          <codeName>ReferenceNumber</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">DIV</property>
              <property name="InnerText" type="string" isRegExp="true">Reference Number: #........-....-....-....-............</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="Button">
          <name>OK</name>
          <codeName>OKButton</codeName>
          <identification>
            <properties>
              <property name="ButtonType" type="string">submit</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">OK</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="WebElement">
          <name>Amount: $1.00</name>
          <codeName>Amount100WebElement</codeName>
          <identification>
            <properties>
              <property name="TagName" type="string">DIV</property>
              <property name="InnerText" type="string">Amount: $1.00</property>
            </properties>
          </identification>
        </object>
      </children>
    </object>
  </objects>
</applicationModel>